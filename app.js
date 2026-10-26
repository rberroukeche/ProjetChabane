const form = document.getElementById("report-form");
const reportType = document.getElementById("report-type");
const zoneField = document.getElementById("zone-field");
const zoneSelect = document.getElementById("zone");
const dealershipField = document.getElementById("dealership-field");
const dealershipSelectField = document.getElementById("dealership-select-field");
const groupFilter = document.getElementById("group-filter");
const groupLabel = document.getElementById("group-label");
const dealership = document.getElementById("dealership");
const previewSection = document.getElementById("preview-section");
const requestStatus = document.getElementById("request-status");
const reportApi = "http://127.0.0.1:8765/report";
const zonesApi = "http://127.0.0.1:8765/zones";
const dealershipsApi = "http://127.0.0.1:8765/dealerships";
let currentPreviewUrl = null;
let dealershipLoadVersion = 0;

function updateDealershipField() {
  const isGroupReport = reportType.value === "group";
  const isDealership = reportType.value === "dealership";
  const needsZone = reportType.value !== "all_zones";
  zoneField.hidden = !needsZone;
  zoneSelect.disabled = !needsZone;
  zoneSelect.required = needsZone;
  dealershipField.hidden = !isGroupReport;
  dealershipSelectField.hidden = !isDealership;
  groupFilter.disabled = !isGroupReport;
  groupFilter.required = isGroupReport;
  dealership.disabled = !isDealership;
  dealership.required = isDealership;
  groupLabel.textContent = "Group to rank";
  groupFilter.options[0].text = isGroupReport
    ? "Choose a group"
    : "All groups";
}

reportType.addEventListener("change", () => {
  updateDealershipField();
  if (reportType.value === "group" || reportType.value === "dealership") {
    loadDealerships();
  } else {
    requestStatus.hidden = true;
  }
});
updateDealershipField();

async function loadDealerships() {
  const requestVersion = ++dealershipLoadVersion;
  groupFilter.disabled = true;
  dealership.disabled = true;
  try {
    const selectedZone = zoneSelect.value || "Zone 4";
    requestStatus.hidden = false;
    requestStatus.classList.remove("error");
    requestStatus.textContent = `Loading dealership and group lists for ${selectedZone}…`;
    const response = await fetch(
      `${dealershipsApi}?zone=${encodeURIComponent(selectedZone)}`
    );
    if (!response.ok) {
      throw new Error("Could not load the dealership list.");
    }
    const responseOptions = await response.json();
    if (requestVersion !== dealershipLoadVersion) {
      return;
    }
    const groupsByKey = new Map();
    const dealershipOptions = responseOptions.map((item) => {
      const name = typeof item === "string" ? item : item.name;
      const suffix = typeof item === "string" ? item.match(/\s+x\s+(.+)$/i) : null;
      const group = typeof item === "string"
        ? (suffix ? suffix[1].trim() : "Group not specified")
        : item.group;
      const groupKey = group.toLocaleLowerCase("fr");
      const displayGroup = groupsByKey.get(groupKey) || group;
      groupsByKey.set(groupKey, displayGroup);
      return { name, group: displayGroup };
    });
    const groups = [...new Set(dealershipOptions.map((item) => item.group))]
      .sort((a, b) => a.localeCompare(b, "fr"));
    groupFilter.replaceChildren(new Option("All groups", ""));
    for (const group of groups) {
      groupFilter.add(new Option(group, group));
    }
    dealership.replaceChildren(new Option("Choose a dealership", ""));
    for (const item of dealershipOptions) {
      dealership.add(new Option(item.name, item.name));
    }
    updateDealershipField();
    requestStatus.hidden = true;
  } catch (error) {
    if (requestVersion !== dealershipLoadVersion) {
      return;
    }
    groupFilter.replaceChildren(new Option("List unavailable", ""));
    dealership.replaceChildren(new Option("List unavailable", ""));
    groupFilter.disabled = true;
    dealership.disabled = true;
    requestStatus.hidden = false;
    requestStatus.classList.add("error");
    requestStatus.textContent = error.message;
  }
}

async function loadZones() {
  requestStatus.hidden = false;
  requestStatus.classList.remove("error");
  requestStatus.textContent = "Loading French zones and Excel data…";
  try {
    const response = await fetch(zonesApi);
    if (!response.ok) {
      throw new Error("Could not load the French zone list.");
    }
    const zones = await response.json();
    zoneSelect.replaceChildren();
    for (const zone of zones) {
      zoneSelect.add(new Option(zone, zone, false, zone === "Zone 4"));
    }
    requestStatus.hidden = true;
  } catch (error) {
    zoneSelect.replaceChildren(new Option("List unavailable", ""));
    requestStatus.hidden = false;
    requestStatus.classList.add("error");
    requestStatus.textContent = error.message;
  }
}

zoneSelect.addEventListener("change", () => {
  if (reportType.value === "group" || reportType.value === "dealership") {
    loadDealerships();
  }
});
loadZones();

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const action = event.submitter.value;
  const submitButton = event.submitter;
  const formData = new FormData(form);
  formData.append("action", action);
  const requestBody = new URLSearchParams(formData);

  requestStatus.hidden = false;
  requestStatus.classList.remove("error");
  requestStatus.textContent = "Generating report…";
  submitButton.disabled = true;

  try {
    const response = await fetch(reportApi, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8" },
      body: requestBody
    });
    if (!response.ok) {
      const errorResponse = await response.json();
      throw new Error(
        errorResponse.error || "The server could not generate the report."
      );
    }

    if (currentPreviewUrl) {
      URL.revokeObjectURL(currentPreviewUrl);
    }
    if (action === "view") {
      const htmlPreview = await response.text();
      if (!response.headers.get("Content-Type")?.includes("text/html")) {
        throw new Error("The server did not return the report preview.");
      }
      const previewBlob = new Blob([htmlPreview], {
        type: "text/html;charset=utf-8"
      });
      currentPreviewUrl = URL.createObjectURL(previewBlob);
      const pdfPreview = document.getElementById("pdf-preview");
      document.getElementById("open-preview").href = currentPreviewUrl;
      pdfPreview.onload = () => {
        const previewDocument = pdfPreview.contentDocument;
        if (previewDocument) {
          pdfPreview.style.height =
            `${Math.max(
              previewDocument.documentElement.scrollHeight,
              previewDocument.body.scrollHeight
            ) + 24}px`;
        }
      };
      pdfPreview.src = currentPreviewUrl;
      previewSection.hidden = false;
      requestStatus.textContent =
        "The report preview is shown below. Use Download PDF to get the PDF file.";
      previewSection.scrollIntoView({ behavior: "smooth", block: "start" });
    } else {
      const pdfBlob = await response.blob();
      if (pdfBlob.type !== "application/pdf") {
        throw new Error("The server did not return a PDF file.");
      }
      currentPreviewUrl = URL.createObjectURL(pdfBlob);
      const filename = response.headers.get("Content-Disposition")
        ?.match(/filename="([^"]+)"/)?.[1] || "report.pdf";
      const downloadLink = document.createElement("a");
      downloadLink.href = currentPreviewUrl;
      downloadLink.download = filename;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      downloadLink.remove();
      requestStatus.textContent = "The report download has started.";
    }
  } catch (error) {
    requestStatus.classList.add("error");
    if (error instanceof TypeError) {
      requestStatus.textContent =
        "Could not connect to the report generator. Start “python web_app.py” " +
        "in the VS Code terminal, then try again.";
    } else {
      requestStatus.textContent = error.message;
    }
  } finally {
    submitButton.disabled = false;
  }
});
