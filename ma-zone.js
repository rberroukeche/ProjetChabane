const apiUrl = "http://127.0.0.1:8765/my-zone-dealerships";
const dealershipList = document.getElementById("dealership-list");
const dealershipCount = document.getElementById("dealership-count");
const zoneTitle = document.getElementById("zone-title");
const requestStatus = document.getElementById("request-status");
const franceMap = document.getElementById("france-map");
const mapStatus = document.getElementById("map-status");
const MAP_WIDTH = 820;
const MAP_HEIGHT = 760;
const MAP_MARGIN = 28;
const REFERENCE_LATITUDE = 46;

function getGeometryPoints(geometry) {
  const points = [];
  function collect(value) {
    if (
      Array.isArray(value) &&
      value.length >= 2 &&
      typeof value[0] === "number" &&
      typeof value[1] === "number"
    ) {
      points.push(value);
    } else if (Array.isArray(value)) {
      for (const child of value) {
        collect(child);
      }
    }
  }
  collect(geometry.coordinates);
  return points;
}

function getGeometryRings(geometry) {
  const polygons = geometry.type === "Polygon"
    ? [geometry.coordinates]
    : geometry.coordinates;
  return polygons.flatMap((polygon) => polygon);
}

function renderFranceMap(geography, dealerships, departmentCodes) {
  const features = geography.features;
  const allPoints = features.flatMap((feature) =>
    getGeometryPoints(feature.geometry)
  );
  const longitudes = allPoints.map((point) => point[0]);
  const latitudes = allPoints.map((point) => point[1]);
  const longitudeScale = Math.cos(REFERENCE_LATITUDE * Math.PI / 180);
  const minX = Math.min(...longitudes) * longitudeScale;
  const maxX = Math.max(...longitudes) * longitudeScale;
  const minY = Math.min(...latitudes);
  const maxY = Math.max(...latitudes);
  const mapScale = Math.min(
    (MAP_WIDTH - MAP_MARGIN * 2) / (maxX - minX),
    (MAP_HEIGHT - MAP_MARGIN * 2) / (maxY - minY)
  );
  const offsetX = (MAP_WIDTH - (maxX - minX) * mapScale) / 2;
  const offsetY = (MAP_HEIGHT - (maxY - minY) * mapScale) / 2;
  const project = (longitude, latitude) => [
    offsetX + (longitude * longitudeScale - minX) * mapScale,
    MAP_HEIGHT - offsetY - (latitude - minY) * mapScale
  ];
  const svgNamespace = "http://www.w3.org/2000/svg";
  franceMap.replaceChildren();

  for (const feature of features) {
    const path = document.createElementNS(svgNamespace, "path");
    const pathData = getGeometryRings(feature.geometry)
      .map((ring) => {
        const [first, ...remaining] = ring;
        const [startX, startY] = project(first[0], first[1]);
        const segments = remaining.map(([longitude, latitude]) => {
          const [x, y] = project(longitude, latitude);
          return `L${x.toFixed(2)},${y.toFixed(2)}`;
        });
        return `M${startX.toFixed(2)},${startY.toFixed(2)}${segments.join("")}Z`;
      })
      .join("");
    path.setAttribute("d", pathData);
    path.setAttribute(
      "class",
      departmentCodes.has(feature.properties.code)
        ? "department department-assigned"
        : "department"
    );
    const title = document.createElementNS(svgNamespace, "title");
    title.textContent = `${feature.properties.name} (${feature.properties.code})`;
    path.appendChild(title);
    franceMap.appendChild(path);
  }

  let plottedCount = 0;
  dealerships.forEach((dealership, index) => {
    if (dealership.longitude === null || dealership.latitude === null) {
      return;
    }
    plottedCount += 1;
    const [x, y] = project(dealership.longitude, dealership.latitude);
    const marker = document.createElementNS(svgNamespace, "g");
    marker.setAttribute("class", "map-marker");
    marker.setAttribute("role", "img");
    marker.setAttribute(
      "aria-label",
      `Map ${index + 1}: ${dealership.name}, ${dealership.records} records`
    );
    const title = document.createElementNS(svgNamespace, "title");
    title.textContent = `${index + 1}. ${dealership.name} — ${dealership.records} records, ${dealership.finished} finished`;
    marker.appendChild(title);
    const circle = document.createElementNS(svgNamespace, "circle");
    circle.setAttribute("cx", x.toFixed(2));
    circle.setAttribute("cy", y.toFixed(2));
    circle.setAttribute("r", "11");
    marker.appendChild(circle);
    const number = document.createElementNS(svgNamespace, "text");
    number.setAttribute("x", x.toFixed(2));
    number.setAttribute("y", (y + 4).toFixed(2));
    number.textContent = String(index + 1);
    marker.appendChild(number);
    franceMap.appendChild(marker);
  });

  const missingCount = dealerships.length - plottedCount;
  mapStatus.hidden = missingCount === 0;
  mapStatus.textContent = missingCount
    ? `${plottedCount} of ${dealerships.length} dealerships located. ${missingCount} location(s) could not be mapped; their Map # is shown as —.`
    : "";
}

function addCell(row, value) {
  const cell = document.createElement("td");
  cell.textContent = value;
  row.appendChild(cell);
}

async function loadAssignedDealerships() {
  try {
    const response = await fetch(apiUrl);
    if (!response.ok) {
      throw new Error("Could not load the assigned dealerships.");
    }
    const data = await response.json();
    zoneTitle.textContent = `My zone — ${data.zone}`;
    dealershipCount.textContent = data.dealerships.length;
    const geographyResponse = await fetch("france-departments.json");
    if (!geographyResponse.ok) {
      throw new Error("Could not load the France department boundaries.");
    }
    const geography = await geographyResponse.json();
    renderFranceMap(
      geography,
      data.dealerships,
      new Set(data.department_codes)
    );
    dealershipList.replaceChildren();

    for (const [index, dealership] of data.dealerships.entries()) {
      const row = document.createElement("tr");
      addCell(
        row,
        dealership.longitude === null ? "—" : String(index + 1)
      );
      addCell(row, dealership.name);
      addCell(row, dealership.records.toLocaleString("en"));
      addCell(row, dealership.finished.toLocaleString("en"));
      addCell(row, dealership.open.toLocaleString("en"));
      addCell(
        row,
        dealership.records
          ? `${dealership.completion_rate.toFixed(1)}%`
          : "No records"
      );
      dealershipList.appendChild(row);
    }

    if (data.dealerships.length === 0) {
      const row = document.createElement("tr");
      addCell(row, "No dealerships are assigned to this zone.");
      row.firstElementChild.colSpan = 6;
      dealershipList.appendChild(row);
    }
  } catch (error) {
    dealershipList.replaceChildren();
    const row = document.createElement("tr");
    addCell(row, "Could not load the assigned dealerships.");
    row.firstElementChild.colSpan = 6;
    dealershipList.appendChild(row);
    requestStatus.hidden = false;
    requestStatus.classList.add("error");
    requestStatus.textContent =
      error instanceof TypeError
        ? "Could not connect to the report generator. Start “python web_app.py” and try again."
        : error.message;
  }
}

loadAssignedDealerships();
