let galaxyData = [];
let pitchAngleChart = null;

// State for sorting, pagination, and filtering
let currentSort = { column: null, direction: 'asc' };
let currentPage = 1;
const pageSize = 100;
let filteredData = [];

// Store last filtered data for download
let lastFilteredData = [];

// Handle file upload
function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  // Display file name
  const fileNameEl = document.getElementById('fileName');
  if (fileNameEl) fileNameEl.textContent = `Selected file: ${file.name}`;

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      // Get the first sheet
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Convert to JSON with raw values
      let rawData = XLSX.utils.sheet_to_json(worksheet, {
        raw: true,
        defval: null, // Use null for empty cells
      });

      // Map Excel headers to internal keys (robust for user columns)
      const headerMap = {
        'id': 'ID',
        'galaxy_name': 'Galaxy_name',
        'j2000_position': 'J2000_position',
        'pa_degrees': 'PA_degrees',
        'pa_err_degrees': 'PA_err_degrees',
        'band_measured_in': 'Band_Measured_In',
        'reference': 'Reference',
        'doi': 'DOI',
        'method': 'Method',
        'band': 'Band',
        'imagesource': 'ImageSource',
        'measurement_type': 'Measurement_Type',
        // Map RA/DEC columns to RA_deg/Dec_deg for cone searches
        'ra': 'RA_deg',
        'dec': 'Dec_deg',
      };

      galaxyData = rawData.map((row) => {
        const mapped = {};
        Object.keys(row).forEach((key) => {
          // Normalize: trim, lower, replace spaces/dashes with underscores
          const normalizedKey = key.trim().toLowerCase().replace(/ |-/g, '_');
          const internalKey = headerMap[normalizedKey] || key;
          mapped[internalKey] = row[key];
        });
        return mapped;
      });

      // Enhanced debug logging
      console.log('First few rows of data:', galaxyData.slice(0, 3));
      console.log('Column names:', Object.keys(galaxyData[0] || {}));
      console.log('Sample pitch angle values:', galaxyData.slice(0, 5).map((g) => g.PA_degrees));
      console.log('Sample error values:', galaxyData.slice(0, 5).map((g) => ({ raw: g.PA_err_degrees, type: typeof g.PA_err_degrees })));
      console.log('Sample Measurement_Type values:', galaxyData.slice(0, 10).map((g) => g.Measurement_Type));

      // Update statistics and display
      updateStatistics();
      displayGalaxies(galaxyData);
      createPitchAngleChart();
      renderAllFilters();
      filterAndDisplay();
      updateFilteredStatistics(galaxyData);
    } catch (e) {
      console.error('File upload error:', e);
      alert('Error processing the file. Please make sure it is a valid Excel file.');
    }
  };

  reader.onerror = function () {
    console.error('Error reading the file');
    alert('Error reading the file. Please try again.');
  };

  reader.readAsArrayBuffer(file);
}

// Update statistics
function updateStatistics() {
  const totalGalaxies = new Set(galaxyData.map((g) => g.Galaxy_name)).size;
  const validPitchAngles = galaxyData
    .map((g) => parseFloat(g.PA_degrees))
    .filter((angle) => !isNaN(angle));

  const avgPitchAngle =
    validPitchAngles.length > 0
      ? (
          validPitchAngles.reduce((sum, angle) => sum + angle, 0) /
          validPitchAngles.length
        ).toFixed(2)
      : 'No data';

  const totalMeasurements = galaxyData.length;

  document.getElementById('totalGalaxies').textContent = totalGalaxies;
  document.getElementById('avgPitchAngle').textContent =
    avgPitchAngle + (avgPitchAngle !== 'No data' ? '°' : '');
  document.getElementById('totalMeasurements').textContent =
    totalMeasurements;
}

// Create pitch angle distribution chart
function createPitchAngleChart(data = null) {
  const sourceData = data || galaxyData;
  const validPitchAngles = sourceData
    .map((g) => parseFloat(g.PA_degrees))
    .filter((angle) => !isNaN(angle));

  // Create histogram data
  const binSize = 5;
  const bins = {};
  validPitchAngles.forEach((angle) => {
    const binIndex = Math.floor(angle / binSize) * binSize;
    bins[binIndex] = (bins[binIndex] || 0) + 1;
  });

  const labels = Object.keys(bins).sort(
    (a, b) => parseFloat(a) - parseFloat(b)
  );
  const dataToRender = labels.map((label) => bins[label]);

  // Destroy existing chart if it exists
  if (pitchAngleChart) {
    pitchAngleChart.destroy();
  }

  // Create new chart
  const ctx = document.getElementById('pitchAngleChart').getContext('2d');
  pitchAngleChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels.map(
        (label) => `${label}° - ${parseFloat(label) + binSize}°`
      ),
      datasets: [
        {
          label: 'Number of Galaxies',
          data: dataToRender,
          backgroundColor: 'rgba(77, 208, 225, 0.5)',
          borderColor: 'rgba(77, 208, 225, 1)',
          borderWidth: 1,
        },
      ],
    },
    options: {
      responsive: true,
      scales: {
        y: {
          beginAtZero: true,
          grid: { color: 'rgba(255, 255, 255, 0.1)' },
          ticks: { color: '#ffffff' },
        },
        x: {
          grid: { color: 'rgba(255, 255, 255, 0.1)' },
          ticks: { color: '#ffffff' },
        },
      },
      plugins: {
        legend: { labels: { color: '#ffffff' } },
      },
    },
  });
}

// Utility: Get all column keys for dynamic sorting
function getAllColumnKeys() {
  if (galaxyData.length === 0) return [];
  return Object.keys(galaxyData[0]);
}

// Utility: Sort data by any column
function sortData(data, column, direction) {
  return [...data].sort((a, b) => {
    const valA = a[column];
    const valB = b[column];

    // Numeric sort if both are numbers
    if (!isNaN(parseFloat(valA)) && !isNaN(parseFloat(valB))) {
      return direction === 'asc'
        ? parseFloat(valA) - parseFloat(valB)
        : parseFloat(valB) - parseFloat(valA);
    }

    // String sort
    return direction === 'asc'
      ? (valA || '').toString().localeCompare((valB || '').toString())
      : (valB || '').toString().localeCompare((valA || '').toString());
  });
}

// Utility: Paginate data
function paginateData(data, page, size) {
  const start = (page - 1) * size;
  return data.slice(start, start + size);
}

// Utility: Download CSV (all data or paged data)
function downloadCSV(data, filename = 'galaxies.csv') {
  if (!data.length) return;
  const keys = Object.keys(data[0]);
  const csvRows = [keys.join(',')];

  data.forEach((row) => {
    csvRows.push(
      keys
        .map((k) => '"' + (row[k] !== undefined && row[k] !== null ? row[k] : '') + '"')
        .join(',')
    );
  });

  const csvContent = csvRows.join('\n');
  const blob = new Blob([csvContent], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Display galaxies in the table with sorting and pagination
function displayGalaxies(galaxies) {
  filteredData = galaxies;
  const keys = getAllColumnKeys();

  // Sort
  let dataToDisplay = galaxies;
  if (currentSort.column) {
    dataToDisplay = sortData(
      dataToDisplay,
      currentSort.column,
      currentSort.direction
    );
  }

  // Paginate
  const totalPages = Math.ceil(dataToDisplay.length / pageSize) || 1;
  if (currentPage > totalPages) currentPage = totalPages;
  const pagedData = paginateData(dataToDisplay, currentPage, pageSize);

  // Table body
  const tbody = document.getElementById('galaxyTableBody');
  tbody.innerHTML = '';
  pagedData.forEach((galaxy) => {
    const row = document.createElement('tr');
    row.innerHTML = keys
      .map((key) => {
        let value = galaxy[key];
        if (key === 'PA_degrees' || key === 'PA_err_degrees') {
          value =
            value === null ||
            value === undefined ||
            value === '' ||
            value === 'N/A'
              ? 'N/A'
              : isNaN(parseFloat(value))
              ? 'N/A'
              : `${parseFloat(value)}°`;
        } else if (
          value === null ||
          value === undefined ||
          value === '' ||
          value === 'N/A'
        ) {
          value = 'N/A';
        }
        return `<td>${value}</td>`;
      })
      .join('');
    tbody.appendChild(row);
  });

  // Table header with sorting
  const thead = document.getElementById('galaxyTableHead');
  thead.innerHTML = '';
  const headerRow = document.createElement('tr');

  keys.forEach((key) => {
    const th = document.createElement('th');
    th.textContent = key;
    th.style.cursor = 'pointer';
    th.onclick = () => {
      if (currentSort.column === key) {
        currentSort.direction =
          currentSort.direction === 'asc' ? 'desc' : 'asc';
      } else {
        currentSort.column = key;
        currentSort.direction = 'asc';
      }
      displayGalaxies(filteredData);
    };
    if (currentSort.column === key) {
      th.textContent += currentSort.direction === 'asc' ? ' ▲' : ' ▼';
    }
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  // Pagination controls
  const paginationDiv = document.getElementById('paginationControls');
  paginationDiv.innerHTML = '';

  const prevBtn = document.createElement('button');
  prevBtn.textContent = 'Prev';
  prevBtn.disabled = currentPage === 1;
  prevBtn.onclick = () => {
    currentPage--;
    displayGalaxies(filteredData);
  };
  paginationDiv.appendChild(prevBtn);

  for (let i = 1; i <= totalPages; i++) {
    const pageBtn = document.createElement('button');
    pageBtn.textContent = i;
    pageBtn.disabled = i === currentPage;
    pageBtn.onclick = () => {
      currentPage = i;
      displayGalaxies(filteredData);
    };
    paginationDiv.appendChild(pageBtn);
  }

  const nextBtn = document.createElement('button');
  nextBtn.textContent = 'Next';
  nextBtn.disabled = currentPage === totalPages;
  nextBtn.onclick = () => {
    currentPage++;
    displayGalaxies(filteredData);
  };
  paginationDiv.appendChild(nextBtn);

  // Download button
  const downloadDiv = document.getElementById('downloadControls');
  downloadDiv.innerHTML = '';
  const downloadBtn = document.createElement('button');
  downloadBtn.textContent = 'Download This Page as CSV';
  downloadBtn.onclick = () => downloadCSV(pagedData);
  downloadDiv.appendChild(downloadBtn);
}

// Render filter dropdowns for Reference, Band, Method, ImageSource, Measurement Type
function renderAllFilters() {
  // Helper function to create checkbox group with check all/uncheck all
  const createCheckboxGroup = (items, prefix, container) => {
    if (!container) return;

    let html = `
      <div class="checkbox-controls mb-2">
        <button class="btn btn-sm btn-outline-info check-all-btn" data-prefix="${prefix}">Check All</button>
        <button class="btn btn-sm btn-outline-info uncheck-all-btn" data-prefix="${prefix}">Uncheck All</button>
      </div>
      <div class="checkbox-group">
    `;

    items.forEach((item) => {
      const safeId = `${prefix}_${String(item).replace(/\s+/g, '_')}`;
      html += `
        <div class="checkbox-item">
          <input type="checkbox" id="${safeId}" class="filter-checkbox" value="${item}">
          <label for="${safeId}">${item}</label>
        </div>
      `;
    });

    html += `</div>`;
    container.innerHTML = html;
  };

  // Reference
  const refContainer = document.getElementById('referenceFilterContainer');
  if (refContainer) {
    const references = Array.from(
      new Set(galaxyData.map((g) => g.Reference).filter(Boolean))
    ).sort();
    createCheckboxGroup(references, 'ref', refContainer);
  }

  // Band
  const bandContainer = document.getElementById('bandFilterContainer');
  if (bandContainer) {
    const bands = Array.from(
      new Set(galaxyData.map((g) => g.Band).filter(Boolean))
    ).sort();
    createCheckboxGroup(bands, 'band', bandContainer);
  }

  // Method
  const methodContainer = document.getElementById('methodFilterContainer');
  if (methodContainer) {
    const methods = Array.from(
      new Set(galaxyData.map((g) => g.Method).filter(Boolean))
    ).sort();
    createCheckboxGroup(methods, 'method', methodContainer);
  }

  // ImageSource
  const imageSourceContainer = document.getElementById('ImageSourceFilterContainer');
  if (imageSourceContainer) {
    const imageSources = Array.from(
      new Set(galaxyData.map((g) => g.ImageSource).filter(Boolean))
    ).sort();
    createCheckboxGroup(imageSources, 'ImageSource', imageSourceContainer);
  }

  // Measurement Type
  const measurementTypeContainer = document.getElementById('measurementTypeFilterContainer');
  if (measurementTypeContainer) {
    const measurementTypes = Array.from(
      new Set(galaxyData.map((g) => g.Measurement_Type).filter(Boolean))
    ).sort();
    createCheckboxGroup(measurementTypes, 'measurementType', measurementTypeContainer);
  }

  // Add event listeners for all checkboxes
  document.querySelectorAll('.filter-checkbox').forEach((checkbox) => {
    checkbox.addEventListener('change', filterAndDisplay);
  });

  // Add event listeners for check all/uncheck all buttons
  document.querySelectorAll('.check-all-btn').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      const prefix = e.target.dataset.prefix;
      document
        .querySelectorAll(`input[type="checkbox"][id^="${prefix}_"]`)
        .forEach((cb) => {
          cb.checked = true;
        });
      filterAndDisplay();
    });
  });

  document.querySelectorAll('.uncheck-all-btn').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      const prefix = e.target.dataset.prefix;
      document
        .querySelectorAll(`input[type="checkbox"][id^="${prefix}_"]`)
        .forEach((cb) => {
          cb.checked = false;
        });
      filterAndDisplay();
    });
  });
}
const ARCSEC_TO_DEG = 1 / 3600;
let coneTargets = [];                        // [{ ra: number, dec: number }, ...]
let coneRadiusDeg = 2.5 * ARCSEC_TO_DEG;     // default 2.5"
let coneOnly = false;                        // if true, show only cone matches

// Parse textarea lines like "25, -30" (degrees)
function parseConeTextarea() {
  const ta = document.getElementById('coneInput');
  if (!ta) return [];
  const lines = ta.value.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const pairs = [];
  for (const line of lines) {
    const parts = line.split(',').map(s => s.trim());
    if (parts.length >= 2) {
      const ra = parseFloat(parts[0]);
      const dec = parseFloat(parts[1]);
      if (isFinite(ra) && isFinite(dec)) pairs.push({ ra, dec });
    }
  }
  return pairs;
}

// Box-match against ANY target (with simple RA wrap handling around 0/360)
function rowMatchesConeBox(row) {
  const ra = parseFloat(row.RA_deg);
  const dec = parseFloat(row.Dec_deg);
  if (!isFinite(ra) || !isFinite(dec)) return false;

  for (const t of coneTargets) {
    // RA wrap-safe delta (assumes RA in [0, 360) or similar)
    let dRA = Math.abs(ra - t.ra);
    if (dRA > 180) dRA = 360 - dRA; // wrap across 0/360 if needed

    const dDec = Math.abs(dec - t.dec);

    if (dRA <= coneRadiusDeg && dDec <= coneRadiusDeg) {
      return true;
    }
  }
  return false;
}
// Update filterAndDisplay to handle multiple selections
function deg2rad(d){ return d * Math.PI / 180; }

// Great-circle separation in degrees (true cone)
function angSepDeg(ra1, dec1, ra2, dec2){
  const r1 = deg2rad(ra1), d1 = deg2rad(dec1);
  const r2 = deg2rad(ra2), d2 = deg2rad(dec2);
  const s = Math.sin((d2 - d1)/2)**2 + Math.cos(d1)*Math.cos(r1)*0 + Math.cos(d1)*Math.cos(d2)*Math.sin((r2 - r1)/2)**2;
  const sin2 = Math.sin((d2 - d1)/2)**2 + Math.cos(d1)*Math.cos(d2)*Math.sin((r2 - r1)/2)**2;
  return (2 * Math.asin(Math.min(1, Math.sqrt(sin2)))) * 180 / Math.PI;
}

function rowMatchesConeCircle(row){
  const ra = parseFloat(row.RA_deg);
  const dec = parseFloat(row.Dec_deg);
  if (!isFinite(ra) || !isFinite(dec)) return false;
  for (const t of coneTargets){
    if (angSepDeg(ra, dec, t.ra, t.dec) <= coneRadiusDeg) return true;
  }
  return false;
}

function filterAndDisplay() {
  let data = [...galaxyData];

  // Get all selected values for each filter type
  const getSelectedValues = (prefix) => {
    return Array.from(
      document.querySelectorAll(`input[type="checkbox"][id^="${prefix}_"]:checked`)
    ).map((cb) => cb.value);
  };

  // Reference filter
  const selectedRefs = getSelectedValues('ref');
  if (selectedRefs.length > 0) {
    data = data.filter((g) => selectedRefs.includes(g.Reference));
  }

  // Band filter
  const selectedBands = getSelectedValues('band');
  if (selectedBands.length > 0) {
    data = data.filter((g) => selectedBands.includes(g.Band));
  }

  // Method filter
  const selectedMethods = getSelectedValues('method');
  if (selectedMethods.length > 0) {
    data = data.filter((g) => selectedMethods.includes(g.Method));
  }

  // ImageSource filter
  const selectedImageSources = getSelectedValues('ImageSource');
  if (selectedImageSources.length > 0) {
    data = data.filter((g) => selectedImageSources.includes(g.ImageSource));
  }

  // Measurement Type filter
  const selectedMeasurementTypes = getSelectedValues('measurementType');
  if (selectedMeasurementTypes.length > 0) {
    data = data.filter((g) => selectedMeasurementTypes.includes(g.Measurement_Type));
  }

  // Search filter
  const searchInput = document.getElementById('searchInput');
  const queryRaw = searchInput ? searchInput.value : '';
  const queryTerms = queryRaw
    .split(/[\n,]+/) // split on commas or newlines
    .map((q) => q.trim().toLowerCase())
    .filter((q) => q.length > 0);

  if (queryTerms.length > 0) {
    data = data.filter((galaxy) => {
      const values = Object.values(galaxy).map((v) => (v || '').toString().toLowerCase());
      return queryTerms.some((term) => values.some((val) => val.includes(term)));
    });
  }

  // Apply cone search constraint (degrees-only, box test) if targets are defined
  if (coneTargets.length > 0) {
   const isMatch = (row) => rowMatchesConeCircle(row);
    // Always narrow to cone matches; this matches the user's expectation that the cone
    // search should refine results rather than merely flagging them.
    data = data.filter(isMatch);
  }

  // Persist the filtered dataset for CSV downloads
  lastFilteredData = data;

  // Reset pagination to the first page before rendering
  currentPage = 1;
  displayGalaxies(data);

  // Update chart based on toggle
  const chartToggle = document.getElementById('chartFilteredToggle');
  if (chartToggle && chartToggle.checked) {
    createPitchAngleChart(data);
  } else {
    createPitchAngleChart(galaxyData);
  }

  // Update filtered statistics
  updateFilteredStatistics(data);
}

// Event listeners
document.addEventListener('DOMContentLoaded', () => {
  // Utility to fetch and process the default Excel file
  fetch('galaxy_database.xlsx')
    .then((response) => response.arrayBuffer())
    .then((arrayBuffer) => {
      try {
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        let rawData = XLSX.utils.sheet_to_json(worksheet, { raw: true, defval: null });

        const headerMap = {
          'id': 'ID',
          'galaxy_name': 'Galaxy_name',
          'j2000_position': 'J2000_position',
          'pa_degrees': 'PA_degrees',
          'pa_err_degrees': 'PA_err_degrees',
          'band_measured_in': 'Band_Measured_In',
          'reference': 'Reference',
          'doi': 'DOI',
          'method': 'Method',
          'band': 'Band',
          'imagesource': 'ImageSource',
          'arm_or_whole': 'Arm_or_Whole',
          'arm': 'Arm_or_Whole',
          'whole': 'Arm_or_Whole',
          'arm/whole': 'Arm_or_Whole',
          'all or whole': 'Arm_or_Whole',
          'measurement_type': 'Measurement_Type',
          // Map RA/DEC columns to RA_deg/Dec_deg for cone searches
          'ra': 'RA_deg',
          'dec': 'Dec_deg',
        };

        galaxyData = rawData.map((row) => {
          const mapped = {};
          Object.keys(row).forEach((key) => {
            const normalizedKey = key.trim().toLowerCase().replace(/ |-/g, '_');
            const internalKey = headerMap[normalizedKey] || key;
            mapped[internalKey] = row[key];
          });
          return mapped;
        });

        updateStatistics();
        displayGalaxies(galaxyData);
        createPitchAngleChart();
        renderAllFilters();
        filterAndDisplay();
        updateFilteredStatistics(galaxyData);
      } catch (e) {
        console.error('Auto-load error:', e);
        alert('Error processing the default galaxy database.');
      }
    })
    .catch((e) => {
      console.error('Failed to fetch galaxy_database.xlsx:', e);
      alert('Could not load the default galaxy database.');
    });

  // Search input with debounce
  const searchInput = document.getElementById('searchInput');
  if (searchInput) {
    let searchTimeout;
    searchInput.addEventListener('input', () => {
      clearTimeout(searchTimeout);
      searchTimeout = setTimeout(() => {
        filterAndDisplay();
      }, 300);
    });
  }

  // Sort select
  const sortSelect = document.getElementById('sortSelect');
  if (sortSelect) {
    sortSelect.addEventListener('change', (e) => {
      sortGalaxies(e.target.value);
    });
  }

  // Add containers if not present
  if (!document.getElementById('paginationControls')) {
    const pagDiv = document.createElement('div');
    pagDiv.id = 'paginationControls';
    pagDiv.style.margin = '10px 0';
    document
      .getElementById('galaxyTable')
      .parentNode.insertBefore(pagDiv, document.getElementById('galaxyTable').nextSibling);
  }
  if (!document.getElementById('downloadControls')) {
    const dlDiv = document.createElement('div');
    dlDiv.id = 'downloadControls';
    dlDiv.style.margin = '10px 0';
    document
      .getElementById('galaxyTable')
      .parentNode.insertBefore(dlDiv, document.getElementById('galaxyTable'));
  }

  // Download all data button
  const downloadAllBtn = document.getElementById('downloadAllBtn');
  if (downloadAllBtn) {
    downloadAllBtn.addEventListener('click', () => {
      downloadCSV(galaxyData, 'galaxies_all.csv');
    });
  }

  // Ensure filters exist
  renderAllFilters();

  // Download filtered data button
  const downloadFilteredBtn = document.getElementById('downloadFilteredBtn');
  if (downloadFilteredBtn) {
    downloadFilteredBtn.addEventListener('click', () => {
      downloadCSV(lastFilteredData, 'galaxies_filtered.csv');
    });
  }

  // Chart toggle
  const chartToggle = document.getElementById('chartFilteredToggle');
  if (chartToggle) {
    chartToggle.addEventListener('change', filterAndDisplay);
  }

  // --- Cone UI wiring ---
  // These listeners connect the cone search inputs to the filtering logic. When the radius changes,
  // update the computed degree radius and re-run filters. When the "Apply" button is clicked, parse
  // the user-specified RA/Dec pairs from the textarea, store them in coneTargets, and apply the
  // filter. The "Clear" button resets the cone filter. The checkbox allows showing only matches.
  const runConeBtn = document.getElementById('runConeSearch');
  const clearConeBtn = document.getElementById('clearConeSearch');
  const coneRadiusInput = document.getElementById('coneRadiusArcsec');
  const coneOnlyChk = document.getElementById('coneOnlyToggle');

  if (coneRadiusInput) {
    coneRadiusInput.addEventListener('change', () => {
      const arcsec = parseFloat(coneRadiusInput.value);
      coneRadiusDeg = (isFinite(arcsec) ? arcsec : 2.5) * ARCSEC_TO_DEG;
      filterAndDisplay();
    });
  }
  if (coneOnlyChk) {
    coneOnlyChk.addEventListener('change', () => {
      coneOnly = coneOnlyChk.checked;
      filterAndDisplay();
    });
  }
  if (runConeBtn) {
    runConeBtn.addEventListener('click', () => {
      coneTargets = parseConeTextarea();
      filterAndDisplay();
    });
  }
  if (clearConeBtn) {
    clearConeBtn.addEventListener('click', () => {
      coneTargets = [];
      const ta = document.getElementById('coneInput');
      if (ta) ta.value = '';
      filterAndDisplay();
    });
  }
});

// Sort functionality
function sortGalaxies(sortBy) {
  if (!sortBy || sortBy === 'none') {
    displayGalaxies(galaxyData);
    return;
  }

  let sortedData = [...galaxyData];

  if (sortBy === 'PA_degrees') {
    sortedData.sort((a, b) => {
      const angleA = parseFloat(a.PA_degrees) || 0;
      const angleB = parseFloat(b.PA_degrees) || 0;
      return angleA - angleB;
    });
  } else if (sortBy === 'PA_err_degrees') {
    sortedData.sort((a, b) => {
      const errorA = parseFloat(a.PA_err_degrees) || 0;
      const errorB = parseFloat(b.PA_err_degrees) || 0;
      return errorA - errorB;
    });
  } else {
    // Generic sort for text fields
    sortedData.sort((a, b) => {
      const valA = (a[sortBy] || '').toString().toLowerCase();
      const valB = (b[sortBy] || '').toString().toLowerCase();
      return valA.localeCompare(valB);
    });
  }

  displayGalaxies(sortedData);
}

function updateFilteredStatistics(filtered) {
  // Unique galaxies
  const totalGalaxies = new Set(filtered.map((g) => g.Galaxy_name)).size;

  // Average pitch angle
  const validPitchAngles = filtered
    .map((g) => parseFloat(g.PA_degrees))
    .filter((angle) => !isNaN(angle));
  const avgPitchAngle =
    validPitchAngles.length > 0
      ? (
          validPitchAngles.reduce((sum, angle) => sum + angle, 0) /
          validPitchAngles.length
        ).toFixed(2)
      : 'No data';

  // Total measurements
  const totalMeasurements = filtered.length;

  document.getElementById('filteredTotalGalaxies').textContent = `Filtered: ${totalGalaxies}`;
  document.getElementById('filteredAvgPitchAngle').textContent =
    `Filtered: ${avgPitchAngle}${avgPitchAngle !== 'No data' ? '°' : ''}`;
  document.getElementById('filteredTotalMeasurements').textContent =
    `Filtered: ${totalMeasurements}`;
}
