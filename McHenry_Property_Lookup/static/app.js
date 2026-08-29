const TABS = ["parcels", "subdivisions", "roads", "addresses"];
let debounceTimers = {};

function $(sel, root = document) { return root.querySelector(sel); }
function $all(sel, root = document) { return Array.from(root.querySelectorAll(sel)); }

function escapeHtml(s) {
  if (s === null || s === undefined) return "";
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function fmtNum(n, digits = 0) {
  if (n === null || n === undefined || n === "") return "";
  const v = Number(n);
  if (Number.isNaN(v)) return "";
  return v.toLocaleString(undefined, { maximumFractionDigits: digits });
}

function initTabs() {
  $all(".tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      $all(".tab").forEach((b) => b.classList.remove("active"));
      $all(".panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      $(`#panel-${btn.dataset.tab}`).classList.add("active");
    });
  });
}

function initSearch(tab, fetchFn, renderFn) {
  const input = $(`#${tab}-q`);
  const resultsEl = $(`#${tab}-results`);
  input.addEventListener("input", () => {
    clearTimeout(debounceTimers[tab]);
    const q = input.value.trim();
    if (!q) {
      resultsEl.innerHTML = "";
      return;
    }
    debounceTimers[tab] = setTimeout(async () => {
      resultsEl.innerHTML = '<div class="result-hint">Searching…</div>';
      try {
        const results = await fetchFn(q);
        renderFn(resultsEl, results);
      } catch (err) {
        resultsEl.innerHTML = `<div class="result-hint">Error: ${escapeHtml(err.message)}</div>`;
      }
    }, 200);
  });
}

async function apiSearch(kind, q) {
  const res = await fetch(`/api/search/${kind}?q=${encodeURIComponent(q)}`);
  if (!res.ok) throw new Error(`search failed (${res.status})`);
  const data = await res.json();
  return data.results || [];
}

function renderTable(el, rows, columns, onRowClick) {
  if (rows.length === 0) {
    el.innerHTML = '<div class="result-hint">No matches.</div>';
    return;
  }
  const thead = `<thead><tr>${columns.map((c) => `<th>${escapeHtml(c.label)}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rows
    .map((row, i) => {
      const cells = columns.map((c) => `<td>${escapeHtml(c.value(row))}</td>`).join("");
      return `<tr class="${onRowClick ? "clickable" : ""}" data-idx="${i}">${cells}</tr>`;
    })
    .join("")}</tbody>`;
  el.innerHTML = `<div class="result-hint">${rows.length} result${rows.length === 1 ? "" : "s"}</div><table class="result-table">${thead}${tbody}</table>`;
  const hint = `${rows.length} shown (results are capped; refine your search for more precise matches)`;
  el.querySelector(".result-hint").textContent = hint;
  if (onRowClick) {
    $all("tbody tr", el).forEach((tr) => {
      tr.addEventListener("click", () => onRowClick(rows[Number(tr.dataset.idx)]));
    });
  }
}

function openModal(html) {
  $("#modal-content").innerHTML = html;
  $("#modal-backdrop").classList.add("open");
}

function closeModal() {
  $("#modal-backdrop").classList.remove("open");
  $("#modal-content").innerHTML = "";
}

function field(label, value) {
  if (value === null || value === undefined || value === "") return "";
  return `<p class="field"><span class="label">${escapeHtml(label)}:</span>${escapeHtml(value)}</p>`;
}

async function openParcelDetail(row) {
  openModal('<div class="result-hint">Loading…</div>');
  try {
    const res = await fetch(`/api/parcels/${encodeURIComponent(row.parcel_number)}`);
    if (!res.ok) throw new Error(`not found (${res.status})`);
    const p = await res.json();
    renderParcelDetail(p);
  } catch (err) {
    openModal(`<div class="result-hint">Error loading parcel: ${escapeHtml(err.message)}</div>`);
  }
}

function normalizeAddress(s) {
  return s.trim().toUpperCase().replace(/\s+/g, " ");
}

function renderParcelDetail(p) {
  const mailingDiffers =
    p.site_address &&
    p.mail_address1 &&
    normalizeAddress(p.site_address) !== normalizeAddress(p.mail_address1);

  const acres = p.parcel_area ? p.parcel_area / 43560 : null;
  const mapLink =
    p.latitude && p.longitude
      ? `<a class="map-link" target="_blank" rel="noopener" href="https://www.google.com/maps?q=${p.latitude},${p.longitude}">View on map (${p.latitude.toFixed(5)}, ${p.longitude.toFixed(5)})</a>`
      : "";

  const landuse = (p.landuse || []).map((l) => l.description).filter(Boolean);
  const landuseText = [...new Set(landuse)].join(", ");

  const districtFields = [
    ["Township", p.township],
    ["Grade School", p.grade_school],
    ["High School", p.high_school],
    ["Unit School", p.unit_school],
    ["Non-High School", p.non_high_school],
    ["Fire District", p.fire_district],
    ["Library District", p.library_district],
    ["Library", p.library],
    ["Park District", p.park_district],
    ["Road District", p.road_district],
    ["Sanitary District", p.sanitary_district],
    ["Drainage", p.drainage],
    ["Forest Preserve", p.forest_preserve],
    ["Community College", p.community_college],
    ["Hospital District", p.hospital_district],
    ["Multi-Township District", p.multi_township_district],
    ["Special District", p.special_district],
    ["Street Light District", p.street_light_district],
    ["TIF District", p.tif_district],
  ].filter(([, v]) => v);

  const html = `
    <div class="detail-header">
      <h2>${escapeHtml(p.owner || "Unknown Owner")}</h2>
      <p class="pin">PIN ${escapeHtml(p.parcel_number)} &middot; ${escapeHtml(p.property_class || "")} &middot; ${escapeHtml(p.tax_status || "")}
        ${mailingDiffers ? '<span class="flag">Mailing address differs from site address</span>' : ""}
      </p>
    </div>
    <div class="detail-grid">
      <div class="detail-section">
        <h3>Site Address</h3>
        ${field("Address", p.site_address)}
        ${field("City / State / ZIP", [p.site_city, p.site_state, p.site_zip].filter(Boolean).join(", "))}
        ${mapLink ? `<p class="field">${mapLink}</p>` : ""}
      </div>
      <div class="detail-section">
        <h3>Mailing Address</h3>
        ${field("Address", [p.mail_address1, p.mail_address2].filter(Boolean).join(" "))}
        ${field("City / State / ZIP", [p.mail_city, p.mail_state, p.mail_zip].filter(Boolean).join(", "))}
      </div>
      <div class="detail-section">
        <h3>Parcel</h3>
        ${field("Area", acres ? `${fmtNum(acres, 2)} acres (${fmtNum(p.parcel_area)} sq ft)` : "")}
        ${field("Land Use", landuseText)}
        ${field("Tax Code", p.tax_code)}
      </div>
      <div class="detail-section">
        <h3>County Tax &amp; Assessment Records</h3>
        <p class="field">This export doesn't include assessed value, sale price, or tax bill amounts &mdash; that lives on the county's own portal.</p>
        <div class="portal-row">
          <code class="pin-copy" id="portal-pin">${escapeHtml(p.county_pin_digits)}</code>
          <button type="button" class="copy-btn" id="copy-pin-btn">Copy PIN</button>
        </div>
        <p class="field">
          <a class="map-link" target="_blank" rel="noopener" href="${escapeHtml(p.county_portal_url)}">Open McHenry County Tax Portal ↗</a>
          &mdash; paste the PIN above into its parcel search.
        </p>
      </div>
      <div class="detail-section">
        <h3>Taxing Districts</h3>
        <div class="districts-grid">
          ${districtFields.map(([label, v]) => field(label, v)).join("")}
        </div>
      </div>
      ${
        p.legal_description
          ? `<div class="detail-section full-width">
              <h3>Legal Description</h3>
              <div class="legal-text">${escapeHtml(p.legal_description)}</div>
            </div>`
          : ""
      }
    </div>
  `;
  openModal(html);

  const copyBtn = $("#copy-pin-btn");
  if (copyBtn) {
    copyBtn.addEventListener("click", async () => {
      const pin = $("#portal-pin").textContent;
      try {
        await navigator.clipboard.writeText(pin);
        copyBtn.textContent = "Copied!";
      } catch (err) {
        copyBtn.textContent = "Copy failed -- select manually";
      }
      setTimeout(() => {
        copyBtn.textContent = "Copy PIN";
      }, 1500);
    });
  }
}

function setup() {
  initTabs();

  initSearch(
    "parcels",
    (q) => apiSearch("parcels", q),
    (el, rows) =>
      renderTable(
        el,
        rows,
        [
          { label: "Owner", value: (r) => r.owner },
          { label: "Site Address", value: (r) => r.site_address },
          { label: "City", value: (r) => r.site_city },
          { label: "PIN", value: (r) => r.parcel_number },
          { label: "Class", value: (r) => r.property_class },
        ],
        openParcelDetail
      )
  );

  initSearch(
    "subdivisions",
    (q) => apiSearch("subdivisions", q),
    (el, rows) =>
      renderTable(el, rows, [
        { label: "Name", value: (r) => r.name },
        { label: "Subcode", value: (r) => r.subcode },
        { label: "Plat Pages", value: (r) => r.pages },
        { label: "Last Update", value: (r) => (r.lastupdate || "").split(" ")[0] },
      ])
  );

  initSearch(
    "roads",
    (q) => apiSearch("roads", q),
    (el, rows) =>
      renderTable(el, rows, [
        { label: "Road Name", value: (r) => r.name },
        { label: "Jurisdiction", value: (r) => r.jurisdiction_name },
        { label: "Class", value: (r) => r.funct_class_name },
        { label: "Route", value: (r) => [r.us_route1, r.state_route1, r.county_route1].filter(Boolean).join(" / ") },
        { label: "Segments", value: (r) => fmtNum(r.segment_count) },
      ])
  );

  initSearch(
    "addresses",
    (q) => apiSearch("addresses", q),
    (el, rows) =>
      renderTable(el, rows, [
        { label: "Address", value: (r) => r.full_address },
        { label: "Municipality", value: (r) => r.municipality },
        { label: "ZIP", value: (r) => r.postal_code },
      ])
  );

  $("#modal-close").addEventListener("click", closeModal);
  $("#modal-backdrop").addEventListener("click", (e) => {
    if (e.target.id === "modal-backdrop") closeModal();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeModal();
  });
}

document.addEventListener("DOMContentLoaded", setup);
