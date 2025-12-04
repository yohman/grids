
    // ===== Constants and globals =====
    const R = 6378137; // WebMercator radius
    const TILE_SIZE = 256;

    let map;
    let gridCells = []; // [{row,col,bbox:[[minLng,minLat],[maxLng,maxLat]]}]
    let paperMode = "A3";
    let paperOrientation = "landscape";
    let gridZoomForExport = null;

    let exportResolutionMode = "high";
    let exportCancelRequested = false;
    let exportTotalCells = 0;

    let tileZoomOverride = null; // ユーザー指定タイルズーム（null の場合は自動）

    let gridLocked = false;
    let isDraggingGrid = false;
    let dragStartMeters = null;
    let dragStartCells = null;

    let currentBaseKey = "esri";

    const basemapConfigs = {
      esri: {
        template:
          "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
        maxZoom: 22,
        maxNativeZoom: 19
      },
      gsi1974: {
        template:
          "https://cyberjapandata.gsi.go.jp/xyz/gazo1/{z}/{x}/{y}.jpg",
        maxZoom: 22,
        maxNativeZoom: 18
      },
      gsiPresent: {
        template:
          "https://cyberjapandata.gsi.go.jp/xyz/seamlessphoto/{z}/{x}/{y}.jpg",
        maxZoom: 22,
        maxNativeZoom: 18
      },
      gsi1961: {
        template:
          "https://cyberjapandata.gsi.go.jp/xyz/ort_old10/{z}/{x}/{y}.png",
        maxZoom: 22,
        maxNativeZoom: 18
      },
      gsi1984: {
        template:
          "https://cyberjapandata.gsi.go.jp/xyz/gazo3/{z}/{x}/{y}.jpg",
        maxZoom: 22,
        maxNativeZoom: 18
      },
      googleSat: {
        template: "https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}",
        maxZoom: 22,
        maxNativeZoom: 21
      },
      googleHybrid: {
        template: "https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}",
        maxZoom: 22,
        maxNativeZoom: 21
      },
      googleMaps: {
        template: "https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}",
        maxZoom: 22,
        maxNativeZoom: 21
      }
    };

    // ===== Helpers =====
    function lngLatToMeters(lng, lat) {
      const x = (lng * Math.PI * R) / 180;
      const y = R * Math.log(Math.tan(Math.PI / 4 + (lat * Math.PI) / 360));
      return { x, y };
    }

    function metersToLngLat(x, y) {
      const lng = (x / R) * (180 / Math.PI);
      const lat =
        (180 / Math.PI) *
        (2 * Math.atan(Math.exp(y / R)) - Math.PI / 2);
      return { lng, lat };
    }

    function latLngToWorldPx(lat, lng, z) {
      const scale = TILE_SIZE * Math.pow(2, z);
      const x = ((lng + 180) / 360) * scale;
      const sinLat = Math.sin((lat * Math.PI) / 180);
      const y =
        (0.5 -
          Math.log((1 + sinLat) / (1 - sinLat)) / (4 * Math.PI)) *
        scale;
      return { x, y };
    }

    function getExportBaseSize() {
      // ベースとなる解像度（1ページあたりのターゲットピクセル数の平方根）
      let base;
      switch (exportResolutionMode) {
        case "low":
          base = 800;
          break;
        case "medium":
          base = 2000;
          break;
        case "high":
          base = 3200;
          break;
        case "veryHigh":
          base = 4800;
          break;
        default:
          base = 3200;
          break;
      }

      // グリッド数が多いときは、PDF 全体のピクセル数が大きくなりすぎて
      // jsPDF が RangeError（Invalid string length）を出すので、
      // 1ページあたりのベースサイズを自動的に下げる。
      const pages = Array.isArray(gridCells) ? gridCells.length : 0;

      // 4セル以下（例: 2×2）は今まで通りフル解像度を維持
      if (!pages || pages <= 4) {
        return base;
      }

      // だいたい 2 億ピクセル程度を上限にする（25ページなどでも安全な範囲）
      const maxTotalPixels = 100_000_000; // 100M
      // このとき 1ページあたりの推奨最大サイズ（1辺の長さ）を逆算
      // sqrt(maxTotalPixels / pages)
      const maxPerPageSide = Math.sqrt(maxTotalPixels / pages);

      // jsPDF の 1ページあたりの最大サイズ（内部制限）とも整合を取る
      const PDF_MAX_SIDE = 14000;
      const safePerPageSide = Math.min(maxPerPageSide, PDF_MAX_SIDE);

      // base が安全な上限を超えている場合のみ縮小
      const adjustedBase = Math.min(base, safePerPageSide);

      // あまり小さくなりすぎると意味がないので、最低 800px は確保
      return Math.max(800, Math.round(adjustedBase));
    }
    function getTileTemplate() {
      const cfg = basemapConfigs[currentBaseKey];
      return cfg ? cfg.template : null;
    }

    function buildTileUrl(template, z, x, y) {
      return template
        .replace("{z}", z)
        .replace("{x}", x)
        .replace("{y}", y);
    }

    // Roughly estimate how many tiles a cell will need at a given zoom
    function estimateTileCountForCell(cell, z) {
      if (!cell || !cell.bbox || !isFinite(z)) return 0;
      const [[minLng, minLat], [maxLng, maxLat]] = cell.bbox;
      const swPx = latLngToWorldPx(minLat, minLng, z);
      const nePx = latLngToWorldPx(maxLat, maxLng, z);

      const minPxX = Math.min(swPx.x, nePx.x);
      const maxPxX = Math.max(swPx.x, nePx.x);
      const minPxY = Math.min(swPx.y, nePx.y);
      const maxPxY = Math.max(swPx.y, nePx.y);

      const maxIndex = Math.pow(2, z) - 1;
      const minTileX = Math.max(0, Math.floor(minPxX / TILE_SIZE));
      const maxTileX = Math.min(maxIndex, Math.floor(maxPxX / TILE_SIZE));
      const minTileY = Math.max(0, Math.floor(minPxY / TILE_SIZE));
      const maxTileY = Math.min(maxIndex, Math.floor(maxPxY / TILE_SIZE));

      const tilesX = maxTileX - minTileX + 1;
      const tilesY = maxTileY - minTileY + 1;
      return tilesX * tilesY;
    }

    function computeAutoTileZoomForCell(cell) {
      if (!cell || !cell.bbox) return 18;
      const [[minLng, minLat], [maxLng, maxLat]] = cell.bbox;

      const base = getExportBaseSize();

      const sw0 = latLngToWorldPx(minLat, minLng, 0);
      const ne0 = latLngToWorldPx(maxLat, maxLng, 0);
      const w0 = Math.abs(ne0.x - sw0.x);
      const h0 = Math.abs(ne0.y - sw0.y);
      const long0 = Math.max(w0, h0);
      if (!isFinite(long0) || long0 <= 0) {
        return 18;
      }

      let autoZ = Math.log2(base / long0);
      if (!isFinite(autoZ)) {
        autoZ = 18;
      }
      autoZ = Math.round(autoZ);

      const cfg = basemapConfigs[currentBaseKey];
      const maxNative = (cfg && typeof cfg.maxNativeZoom === "number")
        ? cfg.maxNativeZoom
        : 22;
      const maxAllowedNative = maxNative + 2;
      const clampedAuto = Math.max(0, Math.min(autoZ, maxAllowedNative));
      return clampedAuto;
    }

    // Per-cell export zoom to roughly match desired pixel size
    function getTileExportZoomForCell(cell) {
      if (!cell || !cell.bbox) {
        return 18;
      }

      const autoZ = computeAutoTileZoomForCell(cell);
      const cfg = basemapConfigs[currentBaseKey];
      let maxNative = (cfg && typeof cfg.maxNativeZoom === "number")
        ? cfg.maxNativeZoom
        : 22;
      const maxAllowedNative = maxNative + 2;

      // ユーザー指定があればそれを優先（ただしネイティブ制限とページ解像度を尊重）
      if (tileZoomOverride !== null && isFinite(tileZoomOverride)) {
        const forced = Math.round(tileZoomOverride);

        // allow up to +2 levels above auto while keeping within native limit
        const allowedBump = autoZ + 2;
        const hardMax = Math.min(maxAllowedNative, allowedBump);
        const zForced = Math.max(0, Math.min(forced, hardMax));
        return zForced;
      }

      // 自動の場合はネイティブ制限の範囲で autoZ をクランプ
      const clampedAuto = Math.max(0, Math.min(autoZ, maxAllowedNative));
      return clampedAuto;
    }

    async function renderCellViaTiles(cell, tileTemplate) {
      const [[minLng, minLat], [maxLng, maxLat]] = cell.bbox;
      const z = getTileExportZoomForCell(cell);

      const swPx = latLngToWorldPx(minLat, minLng, z);
      const nePx = latLngToWorldPx(maxLat, maxLng, z);

      const minPxX = Math.min(swPx.x, nePx.x);
      const maxPxX = Math.max(swPx.x, nePx.x);
      const minPxY = Math.min(swPx.y, nePx.y);
      const maxPxY = Math.max(swPx.y, nePx.y);

      const maxIndex = Math.pow(2, z) - 1;
      const minTileX = Math.max(0, Math.floor(minPxX / TILE_SIZE));
      const maxTileX = Math.min(maxIndex, Math.floor(maxPxX / TILE_SIZE));
      const minTileY = Math.max(0, Math.floor(minPxY / TILE_SIZE));
      const maxTileY = Math.min(maxIndex, Math.floor(maxPxY / TILE_SIZE));

      const tilesX = maxTileX - minTileX + 1;
      const tilesY = maxTileY - minTileY + 1;
      const tileCount = tilesX * tilesY;

      const stitchWidth = tilesX * TILE_SIZE;
      const stitchHeight = tilesY * TILE_SIZE;

      const stitchCanvas = document.createElement("canvas");
      stitchCanvas.width = stitchWidth;
      stitchCanvas.height = stitchHeight;
      const ctx = stitchCanvas.getContext("2d");

      const loadPromises = [];
      for (let ty = minTileY; ty <= maxTileY; ty++) {
        for (let tx = minTileX; tx <= maxTileX; tx++) {
          const dx = (tx - minTileX) * TILE_SIZE;
          const dy = (ty - minTileY) * TILE_SIZE;
          const url = buildTileUrl(tileTemplate, z, tx, ty);
          loadPromises.push(
            new Promise((resolve) => {
              const img = new Image();
              img.crossOrigin = "anonymous";
              img.onload = function () {
                try {
                  ctx.drawImage(img, dx, dy);
                } catch (e) {
                  console.warn("drawImage failed for", url, e);
                }
                resolve(true);
              };
              img.onerror = function () {
                console.warn("tile load failed:", url);
                resolve(false);
              };
              img.src = url;
            })
          );
        }
      }

      await Promise.all(loadPromises);

      const originWorldX = minTileX * TILE_SIZE;
      const originWorldY = minTileY * TILE_SIZE;

      const cropX = Math.round(minPxX - originWorldX);
      const cropY = Math.round(minPxY - originWorldY);
      const cropW = Math.round(maxPxX - minPxX);
      const cropH = Math.round(maxPxY - minPxY);

      const cropCanvas = document.createElement("canvas");
      cropCanvas.width = cropW;
      cropCanvas.height = cropH;
      const cropCtx = cropCanvas.getContext("2d");

      cropCtx.drawImage(
        stitchCanvas,
        cropX,
        cropY,
        cropW,
        cropH,
        0,
        0,
        cropW,
        cropH
      );

      // PDF scaling: clamp by jsPDF limit and by our target base size to avoid huge strings
      const PDF_MAX_SIDE = 14000;
      const targetSide = Math.min(PDF_MAX_SIDE, getExportBaseSize());
      const maxSide = Math.max(cropW, cropH);

      let scale = 1;
      if (maxSide > targetSide) {
        scale = targetSide / maxSide;
      }

      if (scale >= 0.999) {
        // No significant scaling needed; return native resolution
        return { canvas: cropCanvas, width: cropW, height: cropH, tileCount };
      }

      const targetW = Math.max(1, Math.round(cropW * scale));
      const targetH = Math.max(1, Math.round(cropH * scale));

      const scaledCanvas = document.createElement("canvas");
      scaledCanvas.width = targetW;
      scaledCanvas.height = targetH;
      const scaledCtx = scaledCanvas.getContext("2d");
      scaledCtx.drawImage(
        cropCanvas,
        0,
        0,
        cropW,
        cropH,
        0,
        0,
        targetW,
        targetH
      );

      return { canvas: scaledCanvas, width: targetW, height: targetH, tileCount };
    }

    function setStatus(msg) {
      const el = document.getElementById("statusText");
      if (!el) return;
      el.textContent = msg || "";
    }

    function setGridDimensionsText(msg) {
      const el = document.getElementById("gridDims");
      if (!el) return;
      el.textContent = msg || "";
    }

    function setExportOverlay(text, detail, progressFraction) {
      const container = document.getElementById("exportOverlay");
      if (!container) return;
      const textEl = document.getElementById("exportOverlayText");
      const detailEl = document.getElementById("exportOverlayDetail");
      const barOuter = document.getElementById("exportProgressBar");
      const barInner = document.getElementById("exportProgressBarInner");

      if (text && text.length) {
        if (textEl) textEl.textContent = text;
        if (detailEl) detailEl.textContent = detail || "";
        container.style.display = "flex";

        if (typeof progressFraction === "number" && barOuter && barInner) {
          const clamped = Math.max(0, Math.min(1, progressFraction));
          barOuter.style.display = "block";
          barInner.style.width = (clamped * 100).toFixed(1) + "%";
        } else if (barOuter && barInner) {
          barOuter.style.display = "none";
          barInner.style.width = "0%";
        }
      } else {
        container.style.display = "none";
        if (detailEl) detailEl.textContent = "";
        if (barOuter && barInner) {
          barOuter.style.display = "none";
          barInner.style.width = "0%";
        }
      }
    }

    function updateTileZoomOptions() {
      const select = document.getElementById("exportRes");
      if (!select) return;

      const cfg = basemapConfigs[currentBaseKey];
      const maxNative = (cfg && typeof cfg.maxNativeZoom === "number")
        ? cfg.maxNativeZoom
        : 22;
      const maxAllowed = Math.min(22, maxNative + 2);

      let baseZ = map ? Math.round(map.getZoom()) : 12;
      if (gridCells.length) {
        baseZ = computeAutoTileZoomForCell(gridCells[0]);
      }

      const candidates = [];
      candidates.push({ id: "auto", z: baseZ, label: "標準" });
      if (baseZ + 1 <= maxAllowed) {
        candidates.push({ id: "bump1", z: baseZ + 1, label: "細かめ (+1)" });
      }
      if (baseZ + 2 <= maxAllowed) {
        candidates.push({ id: "bump2", z: baseZ + 2, label: "さらに細かく (+2)" });
      }

      const frag = document.createDocumentFragment();
      candidates.forEach((c) => {
        const totalTiles = gridCells.length
          ? gridCells.reduce(
              (sum, cell) => sum + estimateTileCountForCell(cell, c.z),
              0
            )
          : null;
        const tileText =
          totalTiles !== null
            ? `, 約 ${totalTiles.toLocaleString()} タイル`
            : "";
        const opt = document.createElement("option");
        opt.value = c.id;
        opt.dataset.zoom = String(c.z);
        opt.textContent = `${c.label} (z${c.z}${tileText})`;
        frag.appendChild(opt);
      });

      select.innerHTML = "";
      select.appendChild(frag);

      let selectedId = "auto";
      if (tileZoomOverride !== null) {
        const match = candidates.find((c) => c.z === tileZoomOverride);
        if (match) {
          selectedId = match.id;
        }
      }
      select.value = selectedId;
    }

    // ===== URL sync =====
    function updateUrlFromState() {
      if (!map) return;
      const params = new URLSearchParams();

      const center = map.getCenter();
      params.set("lat", center.lat.toFixed(6));
      params.set("lng", center.lng.toFixed(6));
      params.set("zoom", map.getZoom().toFixed(2));

      const basemapSelect = document.getElementById("basemapSelect");
      if (basemapSelect) {
        params.set("basemap", basemapSelect.value);
      }

      const rowsInput = document.getElementById("rows");
      const colsInput = document.getElementById("cols");
      if (rowsInput && colsInput && gridCells.length) {
        params.set("rows", rowsInput.value);
        params.set("cols", colsInput.value);
      }

      if (gridCells.length) {
        let minLng = Infinity, minLat = Infinity;
        let maxLng = -Infinity, maxLat = -Infinity;
        gridCells.forEach((cell) => {
          const [[cMinLng, cMinLat], [cMaxLng, cMaxLat]] = cell.bbox;
          if (cMinLng < minLng) minLng = cMinLng;
          if (cMinLat < minLat) minLat = cMinLat;
          if (cMaxLng > maxLng) maxLng = cMaxLng;
          if (cMaxLat > maxLat) maxLat = cMaxLat;
        });
        if (
          isFinite(minLng) && isFinite(minLat) &&
          isFinite(maxLng) && isFinite(maxLat)
        ) {
          params.set("gswLng", minLng.toFixed(6));
          params.set("gswLat", minLat.toFixed(6));
          params.set("gneLng", maxLng.toFixed(6));
          params.set("gneLat", maxLat.toFixed(6));
        }
        params.set("paper", paperMode);
      }
      params.set("orient", paperOrientation);

      const showBoundariesEl = document.getElementById("showBoundaries");
      if (showBoundariesEl) {
        params.set("boundaries", showBoundariesEl.checked ? "1" : "0");
      }

      const lockEl = document.getElementById("lockGrid");
      if (lockEl) {
        params.set("lock", lockEl.checked ? "1" : "0");
      }

      const zoomInputEl = document.getElementById("zoomLevel");
      if (zoomInputEl) {
        params.set("gridZoom", zoomInputEl.value);
      }

      const qs = params.toString();
      const newUrl = qs
        ? `${window.location.pathname}?${qs}`
        : window.location.pathname;
      window.history.replaceState(null, "", newUrl);
    }

    // ===== Grid drawing & dragging =====
    function getGridBoundsFromCells(cells) {
      if (!cells || !cells.length) return null;
      let minLng = Infinity, minLat = Infinity;
      let maxLng = -Infinity, maxLat = -Infinity;
      cells.forEach((cell) => {
        const [[cMinLng, cMinLat], [cMaxLng, cMaxLat]] = cell.bbox;
        if (cMinLng < minLng) minLng = cMinLng;
        if (cMinLat < minLat) minLat = cMinLat;
        if (cMaxLng > maxLng) maxLng = cMaxLng;
        if (cMaxLat > maxLat) maxLat = cMaxLat;
      });
      if (
        !isFinite(minLng) || !isFinite(minLat) ||
        !isFinite(maxLng) || !isFinite(maxLat)
      ) {
        return null;
      }
      return [
        [minLng, minLat],
        [maxLng, maxLat]
      ];
    }

    function buildGridGeoJSON(cells) {
      return {
        type: "FeatureCollection",
        features: cells.map((cell) => {
          const [[minLng, minLat], [maxLng, maxLat]] = cell.bbox;
          return {
            type: "Feature",
            properties: {
              row: cell.row,
              col: cell.col,
              id: `r${cell.row + 1}_c${cell.col + 1}`
            },
            geometry: {
              type: "Polygon",
              coordinates: [[
                [minLng, minLat],
                [maxLng, minLat],
                [maxLng, maxLat],
                [minLng, maxLat],
                [minLng, minLat]
              ]]
            }
          };
        })
      };
    }

    function ensureGridLayer() {
      if (!map) return;
      const apply = () => {
        const src = map.getSource("grid");
        if (src && src.setData) {
          src.setData(buildGridGeoJSON(gridCells));
        }
      };

      // If the style is not ready yet, wait for load once and then apply.
      if (!map.isStyleLoaded()) {
        map.once("load", apply);
        return;
      }

      apply();
    }

    function pointInAnyCell(lng, lat) {
      for (const cell of gridCells) {
        const [[minLng, minLat], [maxLng, maxLat]] = cell.bbox;
        if (
          lng >= minLng && lng <= maxLng &&
          lat >= minLat && lat <= maxLat
        ) {
          return true;
        }
      }
      return false;
    }

    function onMapMouseDown(e) {
      if (!map || !gridCells.length || gridLocked) return;
      const { lng, lat } = e.lngLat;
      if (!pointInAnyCell(lng, lat)) return;

      const m = lngLatToMeters(lng, lat);
      dragStartMeters = m;
      dragStartCells = gridCells.map((cell) => ({
        row: cell.row,
        col: cell.col,
        bbox: [
          [cell.bbox[0][0], cell.bbox[0][1]],
          [cell.bbox[1][0], cell.bbox[1][1]]
        ]
      }));
      isDraggingGrid = true;

      if (map.dragPan) {
        map.dragPan.disable();
      }
      if (e.originalEvent) {
        e.originalEvent.preventDefault();
        e.originalEvent.stopPropagation();
      }
    }

    function onGridDragMove(e) {
      if (!isDraggingGrid || !map || !dragStartMeters || !dragStartCells) return;
      const { lng, lat } = e.lngLat;
      const currentMeters = lngLatToMeters(lng, lat);
      const dx = currentMeters.x - dragStartMeters.x;
      const dy = currentMeters.y - dragStartMeters.y;

      const newCells = dragStartCells.map((cell) => {
        const [[minLng, minLat], [maxLng, maxLat]] = cell.bbox;
        const minM = lngLatToMeters(minLng, minLat);
        const maxM = lngLatToMeters(maxLng, maxLat);
        const shiftedMin = metersToLngLat(minM.x + dx, minM.y + dy);
        const shiftedMax = metersToLngLat(maxM.x + dx, maxM.y + dy);
        return {
          row: cell.row,
          col: cell.col,
          bbox: [
            [shiftedMin.lng, shiftedMin.lat],
            [shiftedMax.lng, shiftedMax.lat]
          ]
        };
      });

      gridCells = newCells;
      ensureGridLayer();
    }

    function onGridDragEnd() {
      if (!isDraggingGrid) return;
      isDraggingGrid = false;
      dragStartMeters = null;
      dragStartCells = null;
      if (map && map.dragPan) {
        map.dragPan.enable();
      }
      updateUrlFromState();
    }

    // ===== Map initialization =====
    function createMap() {
      const initialCenter = [136.03, 35.35]; // [lng,lat]
      const initialZoom = 11;

      const style = {
        version: 8,
        sources: {
          esri: {
            type: "raster",
            tiles: [basemapConfigs.esri.template],
            tileSize: 256,
            maxzoom: basemapConfigs.esri.maxZoom
          },
          gsi1974: {
            type: "raster",
            tiles: [basemapConfigs.gsi1974.template],
            tileSize: 256,
            maxzoom: basemapConfigs.gsi1974.maxZoom
          },
          gsiPresent: {
            type: "raster",
            tiles: [basemapConfigs.gsiPresent.template],
            tileSize: 256,
            maxzoom: basemapConfigs.gsiPresent.maxZoom
          },
          gsi1961: {
            type: "raster",
            tiles: [basemapConfigs.gsi1961.template],
            tileSize: 256,
            maxzoom: basemapConfigs.gsi1961.maxZoom
          },
          gsi1984: {
            type: "raster",
            tiles: [basemapConfigs.gsi1984.template],
            tileSize: 256,
            maxzoom: basemapConfigs.gsi1984.maxZoom
          },
          googleSat: {
            type: "raster",
            tiles: [basemapConfigs.googleSat.template],
            tileSize: 256,
            maxzoom: basemapConfigs.googleSat.maxZoom
          },
          googleHybrid: {
            type: "raster",
            tiles: [basemapConfigs.googleHybrid.template],
            tileSize: 256,
            maxzoom: basemapConfigs.googleHybrid.maxZoom
          },
          googleMaps: {
            type: "raster",
            tiles: [basemapConfigs.googleMaps.template],
            tileSize: 256,
            maxzoom: basemapConfigs.googleMaps.maxZoom
          },
          grid: {
            type: "geojson",
            data: { type: "FeatureCollection", features: [] }
          },
          overlay: {
            type: "geojson",
            data: { type: "FeatureCollection", features: [] }
          }
        },
        layers: [
          { id: "esri", type: "raster", source: "esri" },
          {
            id: "gsi1974",
            type: "raster",
            source: "gsi1974",
            layout: { visibility: "none" }
          },
          {
            id: "gsiPresent",
            type: "raster",
            source: "gsiPresent",
            layout: { visibility: "none" }
          },
          {
            id: "gsi1961",
            type: "raster",
            source: "gsi1961",
            layout: { visibility: "none" }
          },
          {
            id: "gsi1984",
            type: "raster",
            source: "gsi1984",
            layout: { visibility: "none" }
          },
          {
            id: "googleSat",
            type: "raster",
            source: "googleSat",
            layout: { visibility: "none" }
          },
          {
            id: "googleHybrid",
            type: "raster",
            source: "googleHybrid",
            layout: { visibility: "none" }
          },
          {
            id: "googleMaps",
            type: "raster",
            source: "googleMaps",
            layout: { visibility: "none" }
          },
          {
            id: "grid-fill",
            type: "fill",
            source: "grid",
            paint: {
              "fill-color": "#f97316",
              "fill-opacity": 0.08
            }
          },
          {
            id: "grid-line",
            type: "line",
            source: "grid",
            paint: {
              "line-color": "#ea580c",
              "line-width": 1.2
            }
          },
          {
            id: "overlay-line",
            type: "line",
            source: "overlay",
            paint: {
              "line-color": "#ef4444",
              "line-width": 2
            }
          }
        ]
      };

      map = new maplibregl.Map({
        container: "map",
        style,
        center: initialCenter,
        zoom: initialZoom
      });

      map.addControl(new maplibregl.NavigationControl(), "top-right");

      map.on("load", () => {
        map.on("moveend", updateUrlFromState);
        map.on("zoomend", () => {
          const zInput = document.getElementById("zoomLevel");
          if (zInput) {
            zInput.value = map.getZoom().toFixed(0);
          }
          updateUrlFromState();
        });

        map.on("mousedown", onMapMouseDown);
        map.on("mousemove", onGridDragMove);
        map.on("mouseup", onGridDragEnd);
      });
    }

    function setBasemap(key) {
      if (!map || !basemapConfigs[key]) return;
      currentBaseKey = key;
      const ids = Object.keys(basemapConfigs);
      ids.forEach((id) => {
        const vis = id === key ? "visible" : "none";
        if (map.getLayer(id)) {
          map.setLayoutProperty(id, "visibility", vis);
        }
      });
    }

    // ===== Grid creation =====
    function generateGrid() {
      if (!map) return;

      const rows = Math.max(
        1,
        parseInt(document.getElementById("rows").value, 10) || 1
      );
      const cols = Math.max(
        1,
        parseInt(document.getElementById("cols").value, 10) || 1
      );

      const bounds = map.getBounds();
      const sw = bounds.getSouthWest();
      const ne = bounds.getNorthEast();

      const swM = lngLatToMeters(sw.lng, sw.lat);
      const neM = lngLatToMeters(ne.lng, ne.lat);

      const viewWidth = neM.x - swM.x;
      const viewHeight = neM.y - swM.y;

      let swMAdj = { x: swM.x, y: swM.y };
      let neMAdj = { x: neM.x, y: neM.y };
      let totalW = viewWidth;
      let totalH = viewHeight;

      const centerX = (swM.x + neM.x) / 2;
      const centerY = (swM.y + neM.y) / 2;

      if (paperMode === "A3" || paperMode === "A4") {
        const longOverShort = Math.SQRT2;
        const targetAR =
          paperOrientation === "landscape"
            ? longOverShort
            : 1 / longOverShort;

        const maxCellHeightByHeight = viewHeight / rows;
        const maxCellHeightByWidth = viewWidth / (cols * targetAR);
        const cellHeightMeters = Math.min(
          maxCellHeightByHeight,
          maxCellHeightByWidth
        );
        const cellWidthMeters = cellHeightMeters * targetAR;

        totalW = cellWidthMeters * cols;
        totalH = cellHeightMeters * rows;

        const halfW = totalW / 2;
        const halfH = totalH / 2;
        swMAdj = { x: centerX - halfW, y: centerY - halfH };
        neMAdj = { x: centerX + halfW, y: centerY + halfH };
      }

      const cellW = totalW / cols;
      const cellH = totalH / rows;

      const cells = [];
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const minX = swMAdj.x + c * cellW;
          const maxX = minX + cellW;
          const minY = swMAdj.y + r * cellH;
          const maxY = minY + cellH;

          const swLL = metersToLngLat(minX, minY);
          const neLL = metersToLngLat(maxX, maxY);

          const bbox = [
            [swLL.lng, swLL.lat],
            [neLL.lng, neLL.lat]
          ];
          cells.push({ row: r, col: c, bbox });
        }
      }

      gridCells = cells;
      ensureGridLayer();

      const gridSwLL = metersToLngLat(swMAdj.x, swMAdj.y);
      const gridNeLL = metersToLngLat(neMAdj.x, neMAdj.y);
      const centerLat = (gridSwLL.lat + gridNeLL.lat) / 2;
      const metersPerDegLat = (2 * Math.PI * R) / 360;
      const metersPerDegLng =
        ((2 * Math.PI * R) / 360) * Math.cos((centerLat * Math.PI) / 180);

      const gridWidthKmApprox =
        ((gridNeLL.lng - gridSwLL.lng) * metersPerDegLng) / 1000;
      const gridHeightKmApprox =
        ((gridNeLL.lat - gridSwLL.lat) * metersPerDegLat) / 1000;

      const cellWidthKmApprox = gridWidthKmApprox / cols;
      const cellHeightKmApprox = gridHeightKmApprox / rows;
      const cellAspect =
        cellHeightKmApprox !== 0
          ? cellWidthKmApprox / cellHeightKmApprox
          : NaN;

      const zoomNow = map.getZoom();
      gridZoomForExport = zoomNow;

      const dimsText = `現在のグリッド: ズーム ${zoomNow.toFixed(
        0
      )}, 全体 ≈ ${gridWidthKmApprox.toFixed(
        3
      )} × ${gridHeightKmApprox.toFixed(
        3
      )} km; 1セル ≈ ${cellWidthKmApprox.toFixed(
        3
      )} × ${cellHeightKmApprox.toFixed(
        3
      )} km (縦横比 ${
        isFinite(cellAspect) ? cellAspect.toFixed(3) : "–"
      }, 用紙モード ${paperMode} / ${
        paperOrientation === "landscape" ? "横" : "縦"
      }).`;

      setGridDimensionsText(dimsText);
      setStatus(`グリッドを作成しました: ${rows} × ${cols}。${dimsText}`);

      const zoomInputEl = document.getElementById("zoomLevel");
      if (zoomInputEl) {
        zoomInputEl.value = zoomNow.toFixed(0);
      }

      updateTileZoomOptions();
      updateUrlFromState();
    }

    // ===== Search =====
    function searchLocation() {
      if (!map) return;
      const q = document.getElementById("searchInput").value.trim();
      if (!q) return;
      setStatus(`「${q}」を検索中です…`);
      fetch(
        "https://nominatim.openstreetmap.org/search?format=json&q=" +
          encodeURIComponent(q)
      )
        .then((res) => res.json())
        .then((data) => {
          if (!data || !data.length) {
            setStatus(`「${q}」は見つかりませんでした。`);
            return;
          }
          const best = data[0];
          const lat = parseFloat(best.lat);
          const lon = parseFloat(best.lon);
          map.flyTo({ center: [lon, lat], zoom: 12 });
          setStatus("場所を移動しました: " + best.display_name);
        })
        .catch((err) => {
          console.error(err);
          setStatus("検索に失敗しました。もう一度お試しください。");
        });
    }

    // ===== GeoJSON overlay =====
    function loadGeoJSONFromUrl(presetUrl) {
      if (!map) return;
      const input = document.getElementById("geojsonUrl");
      if (!input) return;
      const url = (presetUrl || input.value || "").trim();
      if (!url) {
        setStatus("GeoJSON の URL を入力してください。");
        return;
      }
      // Keep the input synchronized when a preset is used so it can be copied.
      if (presetUrl) {
        input.value = presetUrl;
      }
      setStatus("GeoJSON を読み込み中です...");
      fetch(url)
        .then((res) => {
          if (!res.ok) throw new Error("HTTP " + res.status);
          return res.json();
        })
        .then((data) => {
          const src = map.getSource("overlay");
          if (src && src.setData) {
            src.setData(data);
          }
          // Simple fit bounds
          try {
            let minLng = Infinity, minLat = Infinity;
            let maxLng = -Infinity, maxLat = -Infinity;
            const walk = (geom) => {
              if (!geom) return;
              const type = geom.type;
              if (type === "Point") {
                const [lng, lat] = geom.coordinates;
                if (lng < minLng) minLng = lng;
                if (lat < minLat) minLat = lat;
                if (lng > maxLng) maxLng = lng;
                if (lat > maxLat) maxLat = lat;
              } else if (type === "LineString" || type === "MultiPoint") {
                geom.coordinates.forEach(([lng, lat]) => {
                  if (lng < minLng) minLng = lng;
                  if (lat < minLat) minLat = lat;
                  if (lng > maxLng) maxLng = lng;
                  if (lat > maxLat) maxLat = lat;
                });
              } else if (type === "Polygon" || type === "MultiLineString") {
                geom.coordinates.forEach((ring) =>
                  ring.forEach(([lng, lat]) => {
                    if (lng < minLng) minLng = lng;
                    if (lat < minLat) minLat = lat;
                    if (lng > maxLng) maxLng = lng;
                    if (lat > maxLat) maxLat = lat;
                  })
                );
              } else if (type === "MultiPolygon") {
                geom.coordinates.forEach((poly) =>
                  poly.forEach((ring) =>
                    ring.forEach(([lng, lat]) => {
                      if (lng < minLng) minLng = lng;
                      if (lat < minLat) minLat = lat;
                      if (lng > maxLng) maxLng = lng;
                      if (lat > maxLat) maxLat = lat;
                    })
                  )
                );
              } else if (type === "GeometryCollection") {
                geom.geometries.forEach(walk);
              }
            };

            if (data.type === "FeatureCollection") {
              data.features.forEach((f) => walk(f.geometry));
            } else if (data.type === "Feature") {
              walk(data.geometry);
            } else {
              walk(data);
            }

            if (
              isFinite(minLng) && isFinite(minLat) &&
              isFinite(maxLng) && isFinite(maxLat)
            ) {
              map.fitBounds(
                [
                  [minLng, minLat],
                  [maxLng, maxLat]
                ],
                { padding: 20, duration: 800 }
              );
            }
          } catch (e) {
            console.warn("Could not fit bounds for GeoJSON overlay:", e);
          }
          setStatus("GeoJSON レイヤーを読み込みました。");
        })
        .catch((err) => {
          console.error(err);
          setStatus("GeoJSON の読み込みに失敗しました。CORS や URL を確認してください。");
        });
    }

    function clearGeoJSONOverlay() {
      if (!map) return;
      const src = map.getSource("overlay");
      if (src && src.setData) {
        src.setData({ type: "FeatureCollection", features: [] });
      }
      setStatus("GeoJSON レイヤーをクリアしました。");
    }

    // ===== Export current view (PNG) via MapLibre canvas =====
    function exportCurrentView() {
      if (!map) return;
      const showBoundariesEl = document.getElementById("showBoundaries");
      const showBoundaries = !showBoundariesEl || showBoundariesEl.checked;

      const hideGrid = !showBoundaries;
      if (hideGrid && map.getLayer("grid-line") && map.getLayer("grid-fill")) {
        map.setLayoutProperty("grid-line", "visibility", "none");
        map.setLayoutProperty("grid-fill", "visibility", "none");
      }

      try {
        const canvas = map.getCanvas();
        const dataURL = canvas.toDataURL("image/png");
        const link = document.createElement("a");
        const ts = new Date().toISOString().replace(/[:.]/g, "-");
        link.href = dataURL;
        link.download = `map_view_${ts}.png`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setStatus("現在の表示を PNG として保存しました。");
      } catch (e) {
        console.error(e);
        setStatus("エクスポートに失敗しました（ブラウザの権限を確認してください）。");
      } finally {
        if (hideGrid && map.getLayer("grid-line") && map.getLayer("grid-fill")) {
          map.setLayoutProperty("grid-line", "visibility", "visible");
          map.setLayoutProperty("grid-fill", "visibility", "visible");
        }
      }
    }

    // ===== Export all cells to one PDF (tile-based, native resolution) =====
    async function exportAllCells() {
      if (!map) return;
      if (!gridCells.length) {
        alert("先にグリッドを作成してください。");
        return;
      }

      const tileTemplate = getTileTemplate();
      if (!tileTemplate) {
        setStatus("現在のベースマップはタイルエクスポートに対応していません。別のベースマップを試してください。");
        return;
      }

      exportCancelRequested = false;

      const exportBtn = document.getElementById("exportBtn");
      const exportOneBtn = document.getElementById("exportOneBtn");
      exportBtn.disabled = true;
      exportOneBtn.disabled = true;

      const { jsPDF } = window.jspdf;
      let pdf = null;

      const total = gridCells.length;
      exportTotalCells = total;
      let totalTilesSoFar = 0;

      const showBoundariesEl = document.getElementById("showBoundaries");
      const showBoundaries = !showBoundariesEl || showBoundariesEl.checked;

      setStatus(`セル 1 / ${total} を出力中...`);
      setExportOverlay(`セル 1 / ${total} を出力中...`, "", 0);

      function finalizeExport({ save, message }) {
        if (save && pdf) {
          pdf.save("map_grid.pdf");
        }
        exportTotalCells = 0;
        exportBtn.disabled = false;
        exportOneBtn.disabled = false;
        setStatus(message);
        setExportOverlay("");
      }

      try {
        for (let index = 0; index < total; index++) {
          if (exportCancelRequested) {
            finalizeExport({
              save: false,
              message: "エクスポートをキャンセルしました。"
            });
            return;
          }

          const cell = gridCells[index];
          const stepText = `セル ${index + 1} / ${total} を出力中...`;
          setStatus(stepText);

          const {
            canvas: cropCanvas,
            width: cropW,
            height: cropH,
            tileCount
          } = await renderCellViaTiles(cell, tileTemplate);

          totalTilesSoFar += (tileCount || 0);
          const detailText = `このセル: 約 ${tileCount || 0} タイル / 累計: 約 ${totalTilesSoFar} タイル`;
          const progress = total > 0 ? (index + 1) / total : 0;
          setExportOverlay(stepText, detailText, progress);

          // (No boundaries drawn)

          const dataURL = cropCanvas.toDataURL("image/png");

          if (!pdf) {
            const orientation = cropW >= cropH ? "landscape" : "portrait";
            pdf = new jsPDF({
              orientation,
              unit: "px",
              format: [cropW, cropH]
            });
          } else {
            pdf.addPage(
              [cropW, cropH],
              cropW >= cropH ? "landscape" : "portrait"
            );
          }

          const pageWidth = pdf.internal.pageSize.getWidth();
          const pageHeight = pdf.internal.pageSize.getHeight();
          pdf.addImage(dataURL, "PNG", 0, 0, pageWidth, pageHeight);
        }

        finalizeExport({
          save: true,
          message: `エクスポート完了: ${total} セルを PDF に保存しました。`
        });
      } catch (e) {
        console.error(e);
        finalizeExport({
          save: !!pdf,
          message: "タイルエクスポート中にエラーが発生しました。"
        });
      }
    }

    // ===== Export overview / index page with padding =====
    async function exportIndexPage() {
      if (!map) return;
      if (!gridCells.length) {
        alert("先にグリッドを作成してください。");
        return;
      }

      const bounds = getGridBoundsFromCells(gridCells);
      if (!bounds) {
        setStatus("グリッドの範囲が取得できませんでした。");
        return;
      }

      const originalView = {
        center: map.getCenter(),
        zoom: map.getZoom(),
        bearing: map.getBearing ? map.getBearing() : 0,
        pitch: map.getPitch ? map.getPitch() : 0
      };

      const gridLineVis = map.getLayoutProperty("grid-line", "visibility") || "visible";
      const gridFillVis = map.getLayoutProperty("grid-fill", "visibility") || "visible";
      // Force grid lines visible for index print
      map.setLayoutProperty("grid-line", "visibility", "visible");
      map.setLayoutProperty("grid-fill", "visibility", "visible");

      const waitForIdle = () =>
        new Promise((resolve) => map.once("idle", () => setTimeout(resolve, 120)));

      const paddingPx = 60; // add breathing room around grid edges
      map.fitBounds(bounds, { padding: paddingPx, duration: 0 });
      await waitForIdle();

      try {
        const canvas = map.getCanvas();
        const dataURL = canvas.toDataURL("image/png");
        const win = window.open("", "_blank");
        if (win) {
          const printHtml =
            "<!doctype html>" +
            "<html>" +
            "<head><title>Grid Index</title></head>" +
            '<body style="margin:24px; display:flex; justify-content:center; align-items:center; background:#f3f4f6;">' +
            '<img src="' + dataURL + '" style="max-width:100%; height:auto; border:1px solid #d1d5db; padding:8px; background:#fff;" />' +
            "</body>" +
            "</html>";
          win.document.write(printHtml);
          win.document.close();
          win.focus();
          win.print();
        } else {
          // Fallback: download PNG if popup blocked
          const link = document.createElement("a");
          link.href = dataURL;
          link.download = "grid_index.png";
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
        }
        setStatus("インデックスページを印刷/保存用に開きました。");
      } catch (e) {
        console.error(e);
        setStatus("インデックス出力に失敗しました。");
      } finally {
        map.setLayoutProperty("grid-line", "visibility", gridLineVis);
        map.setLayoutProperty("grid-fill", "visibility", gridFillVis);
        map.jumpTo({
          center: originalView.center,
          zoom: originalView.zoom,
          bearing: originalView.bearing,
          pitch: originalView.pitch
        });
      }
    }

    // ===== DOMContentLoaded wiring =====
    window.addEventListener("DOMContentLoaded", () => {
      const params = new URLSearchParams(window.location.search);

      createMap();

      const zoomInputEl = document.getElementById("zoomLevel");
      if (zoomInputEl && map) {
        zoomInputEl.value = map.getZoom().toFixed(0);
      }

      const lockEl = document.getElementById("lockGrid");
      if (lockEl) {
        if (params.has("lock")) {
          lockEl.checked = params.get("lock") === "1";
        }
        gridLocked = lockEl.checked;
        lockEl.addEventListener("change", () => {
          gridLocked = lockEl.checked;
          updateUrlFromState();
        });
      }

      const showBoundariesEl = document.getElementById("showBoundaries");
      if (showBoundariesEl) {
        if (params.has("boundaries")) {
          showBoundariesEl.checked = params.get("boundaries") !== "0";
        }
        showBoundariesEl.addEventListener("change", () => {
          updateUrlFromState();
        });
      }

      const basemapSelect = document.getElementById("basemapSelect");
      if (basemapSelect && params.has("basemap")) {
        const key = params.get("basemap");
        if (key && basemapConfigs[key]) {
          basemapSelect.value = key;
          setBasemap(key);
        }
      }

      if (params.has("lat") && params.has("lng")) {
        const lat = parseFloat(params.get("lat"));
        const lng = parseFloat(params.get("lng"));
        if (isFinite(lat) && isFinite(lng)) {
          let zoom = map.getZoom();
          if (params.has("zoom")) {
            const zParam = parseFloat(params.get("zoom"));
            if (isFinite(zParam)) zoom = zParam;
          }
          map.jumpTo({ center: [lng, lat], zoom });
          if (zoomInputEl) zoomInputEl.value = zoom.toFixed(0);
        }
      } else if (params.has("zoom")) {
        const zParam = parseFloat(params.get("zoom"));
        if (isFinite(zParam)) {
          const c = map.getCenter();
          map.jumpTo({ center: [c.lng, c.lat], zoom: zParam });
          if (zoomInputEl) zoomInputEl.value = zParam.toFixed(0);
        }
      }

      const rowsInput = document.getElementById("rows");
      const colsInput = document.getElementById("cols");
      if (rowsInput && params.has("rows")) {
        rowsInput.value = params.get("rows");
      }
      if (colsInput && params.has("cols")) {
        colsInput.value = params.get("cols");
      }
      if (zoomInputEl && params.has("gridZoom")) {
        zoomInputEl.value = params.get("gridZoom");
        const gz = parseFloat(params.get("gridZoom"));
        if (isFinite(gz)) {
          gridZoomForExport = gz;
        }
      }

      if (params.has("paper")) {
        const p = params.get("paper");
        if (p === "A3" || p === "A4" || p === "custom") {
          paperMode = p;
        }
      }
      if (params.has("orient")) {
        const o = params.get("orient");
        if (o === "landscape" || o === "portrait") {
          paperOrientation = o;
        }
      }

      // Restore grid from URL if present
      if (
        params.has("rows") &&
        params.has("cols") &&
        params.has("gswLng") &&
        params.has("gswLat") &&
        params.has("gneLng") &&
        params.has("gneLat")
      ) {
        const rows = Math.max(1, parseInt(params.get("rows"), 10) || 1);
        const cols = Math.max(1, parseInt(params.get("cols"), 10) || 1);
        const swLng = parseFloat(params.get("gswLng"));
        const swLat = parseFloat(params.get("gswLat"));
        const neLng = parseFloat(params.get("gneLng"));
        const neLat = parseFloat(params.get("gneLat"));

        if (
          isFinite(swLng) && isFinite(swLat) &&
          isFinite(neLng) && isFinite(neLat)
        ) {
          const swM = lngLatToMeters(swLng, swLat);
          const neM = lngLatToMeters(neLng, neLat);
          const totalW = neM.x - swM.x;
          const totalH = neM.y - swM.y;
          const cellW = totalW / cols;
          const cellH = totalH / rows;

          const cells = [];
          for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
              const minX = swM.x + c * cellW;
              const maxX = minX + cellW;
              const minY = swM.y + r * cellH;
              const maxY = minY + cellH;

              const swLL = metersToLngLat(minX, minY);
              const neLL = metersToLngLat(maxX, maxY);

              const bbox = [
                [swLL.lng, swLL.lat],
                [neLL.lng, neLL.lat]
              ];
              cells.push({ row: r, col: c, bbox });
            }
          }

          gridCells = cells;
          ensureGridLayer();

          const gridSwLL = { lng: swLng, lat: swLat };
          const gridNeLL = { lng: neLng, lat: neLat };
          const centerLat = (gridSwLL.lat + gridNeLL.lat) / 2;
          const metersPerDegLat = (2 * Math.PI * R) / 360;
          const metersPerDegLng =
            ((2 * Math.PI * R) / 360) * Math.cos((centerLat * Math.PI) / 180);

          const gridWidthKmApprox =
            ((gridNeLL.lng - gridSwLL.lng) * metersPerDegLng) / 1000;
          const gridHeightKmApprox =
            ((gridNeLL.lat - gridSwLL.lat) * metersPerDegLat) / 1000;

          const cellWidthKmApprox = gridWidthKmApprox / cols;
          const cellHeightKmApprox = gridHeightKmApprox / rows;
          const cellAspect =
            cellHeightKmApprox !== 0
              ? cellWidthKmApprox / cellHeightKmApprox
              : NaN;

          const zoomNow = map.getZoom();
          gridZoomForExport = zoomNow;

          const dimsText = `現在のグリッド: ズーム ${zoomNow.toFixed(
            0
          )}, 全体 ≈ ${gridWidthKmApprox.toFixed(
            3
          )} × ${gridHeightKmApprox.toFixed(
            3
          )} km; 1セル ≈ ${cellWidthKmApprox.toFixed(
            3
          )} × ${cellHeightKmApprox.toFixed(
            3
          )} km (縦横比 ${
              isFinite(cellAspect) ? cellAspect.toFixed(3) : "–"
            }, 用紙モード ${paperMode} / ${
              paperOrientation === "landscape" ? "横" : "縦"
            }).`;

          setGridDimensionsText(dimsText);
          setStatus(`URL からグリッドを復元しました。${dimsText}`);
          updateTileZoomOptions();

          if (!(params.has("lat") && params.has("lng"))) {
            map.fitBounds(
              [
                [swLng, swLat],
                [neLng, neLat]
              ],
              { padding: 20, duration: 0 }
            );
          }
        }
      }

      // Paper preset buttons
      const paperButtons = Array.from(document.querySelectorAll(".paper-btn"));
      function updatePaperButtons() {
        paperButtons.forEach((btn) => {
          const mode = btn.getAttribute("data-paper");
          if (mode === paperMode) {
            btn.classList.add("paper-active");
          } else {
            btn.classList.remove("paper-active");
          }
        });
      }
      paperButtons.forEach((btn) => {
        btn.addEventListener("click", () => {
          const mode = btn.getAttribute("data-paper");
          if (!mode) return;
          paperMode = mode;
          updatePaperButtons();
          setStatus(
            `用紙モードを ${paperMode} に変更しました。グリッドを再生成してください。`
          );
          updateUrlFromState();
        });
      });
      updatePaperButtons();

      // Orientation buttons
      const orientButtons = Array.from(document.querySelectorAll(".orient-btn"));
      function updateOrientButtons() {
        orientButtons.forEach((btn) => {
          const o = btn.getAttribute("data-orient");
          if (o === paperOrientation) {
            btn.classList.add("orient-active");
          } else {
            btn.classList.remove("orient-active");
          }
        });
      }
      orientButtons.forEach((btn) => {
        btn.addEventListener("click", () => {
          const o = btn.getAttribute("data-orient");
          if (!o) return;
          if (o === "landscape" || o === "portrait") {
            paperOrientation = o;
            updateOrientButtons();
            setStatus(
              `用紙の向きを ${
                paperOrientation === "landscape" ? "横" : "縦"
              } に変更しました。グリッドを再生成してください。`
            );
            updateUrlFromState();
          }
        });
      });
      updateOrientButtons();

      // Buttons / controls
      document
        .getElementById("generateGridBtn")
        .addEventListener("click", generateGrid);

      document
        .getElementById("basemapSelect")
        .addEventListener("change", (e) => {
          setBasemap(e.target.value);
          updateTileZoomOptions();
          updateUrlFromState();
        });

      document
        .getElementById("exportBtn")
        .addEventListener("click", exportAllCells);

      document
        .getElementById("exportOneBtn")
        .addEventListener("click", exportCurrentView);

      document
        .getElementById("printIndexBtn")
        .addEventListener("click", exportIndexPage);

      document
        .getElementById("searchBtn")
        .addEventListener("click", searchLocation);

      document
        .getElementById("searchInput")
        .addEventListener("keydown", (e) => {
          if (e.key === "Enter") searchLocation();
        });

      const loadGeojsonBtn = document.getElementById("loadGeojsonBtn");
      if (loadGeojsonBtn) {
        loadGeojsonBtn.addEventListener("click", loadGeoJSONFromUrl);
      }
      const clearGeojsonBtn = document.getElementById("clearGeojsonBtn");
      if (clearGeojsonBtn) {
        clearGeojsonBtn.addEventListener("click", clearGeoJSONOverlay);
      }

      // Sample Takashima GeoJSON helper
      const takashimaLink = document.getElementById("takashimaLink");
      if (takashimaLink) {
        takashimaLink.addEventListener("click", (e) => {
          // Keep navigation available in new tab while also loading on the map.
          e.preventDefault();
          const sampleUrl = takashimaLink.getAttribute("href");
          if (sampleUrl) {
            loadGeoJSONFromUrl(sampleUrl);
          }
        });
      }

      const exportResSelect = document.getElementById("exportRes");
      if (exportResSelect) {
        updateTileZoomOptions();
        exportResSelect.addEventListener("change", (e) => {
          const selected = e.target.selectedOptions
            ? e.target.selectedOptions[0]
            : null;
          if (!selected) return;

          if (selected.value === "auto") {
            tileZoomOverride = null;
            setStatus("タイルズーム: 自動計算 (推奨) にしました。");
            return;
          }

          const zStr = selected.dataset.zoom;
          const zVal = parseInt(zStr, 10);
          if (isFinite(zVal)) {
            tileZoomOverride = zVal;
            setStatus(`タイルズーム: z${zVal} に固定します。`);
          }
        });
      }

      const cancelBtn = document.getElementById("cancelExportBtn");
      if (cancelBtn) {
        cancelBtn.addEventListener("click", () => {
          if (!exportCancelRequested) {
            exportCancelRequested = true;
            setStatus("エクスポートをキャンセルしています...");
          }
        });
      }

      updateUrlFromState();
    });
  