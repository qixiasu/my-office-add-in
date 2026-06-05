/**
 * Parse delimiter-separated text into a 2D array.
 * Follows RFC 4180: handles quoted fields, embedded delimiters,
 * embedded newlines, and doubled quotes.
 * Strips UTF-8 BOM if present.
 *
 * NOTE: This loads the full array into memory. For large files, use
 * scanCSVFile / parseCSVFile instead — they stream the file in chunks.
 *
 * @param {string} text - Raw file content
 * @param {string} delimiter - Field delimiter (default ",")
 * @returns {string[][]}
 */
function parseCSV(text, delimiter) {
  if (!text || text.length === 0) {
    return [];
  }

  var sep = delimiter || ",";

  // Strip UTF-8 BOM
  if (text.charCodeAt(0) === 0xfeff) {
    text = text.slice(1);
  }

  var rows = [];
  var row = [];
  var field = "";
  var inQuotes = false;

  for (var i = 0; i < text.length; i++) {
    var ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        var nextCh = text[i + 1];
        if (nextCh === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
    } else {
      if (ch === '"' && field === "") {
        inQuotes = true;
      } else if (ch === sep) {
        row.push(field);
        field = "";
      } else if (ch === "\n") {
        row.push(field);
        field = "";
        if (row.length > 0 || rows.length > 0) {
          rows.push(row);
        }
        row = [];
      } else if (ch === "\r") {
        if (text[i + 1] === "\n") {
          i++;
        }
        row.push(field);
        field = "";
        if (row.length > 0 || rows.length > 0) {
          rows.push(row);
        }
        row = [];
      } else {
        field += ch;
      }
    }
  }

  // Last field
  row.push(field);
  if (row.length > 1 || row[0] !== "" || rows.length === 0) {
    rows.push(row);
  }

  return rows;
}

// ── Chunked file reading ──────────────────────────────────────────────

var FILE_CHUNK_SIZE = 1 * 1024 * 1024; // 1 MB per read chunk

/**
 * Read a File in binary chunks, decode as text, and feed each text chunk
 * to onChunk. Yields to the event loop between chunks so the UI stays
 * responsive and the browser never loads the full file into a single string.
 *
 * Uses two strategies:
 *   UTF‑8   → TextDecoder with {stream:true} (fast, handles boundaries natively)
 *   non‑UTF → manual carry‑over at ArrayBuffer level, then TextDecoder per chunk
 *             (avoids Chromium bugs with stream mode on non‑UTF‑8 decoders)
 *
 * @param {File} file     - Browser File object
 * @param {string} encoding - Text encoding ("utf-8", "gbk", "gb18030", "big5")
 * @param {function} onChunk - Called with (text, bytesRead, totalBytes) each chunk
 * @param {function} onDone  - Called when complete
 * @param {function} onError - Called on read error
 */
function readFileInChunks(file, encoding, onChunk, onDone, onError) {
  var enc = encoding || "utf-8";
  var isUTF = enc === "utf-8" || enc === "utf8";

  // ── UTF‑8 fast path ─────────────────────────────────────────────────
  if (isUTF) {
    return readFileInChunksUTF8(file, onChunk, onDone, onError);
  }

  // ── Non‑UTF‑8 path (GBK / Big5 / GB18030) ──────────────────────────
  var offset = 0;
  var decoder = createTextDecoder(enc);
  var carryover = new Uint8Array(0);
  var firstChunk = true;

  function next() {
    if (offset >= file.size) {
      // Decode any leftover bytes
      if (carryover.length > 0) {
        try {
          var tail = decoder.decode(carryover);
          if (tail) onChunk(tail, offset, file.size);
        } catch (ignored) {}
      }
      onDone();
      return;
    }

    var end = Math.min(offset + FILE_CHUNK_SIZE, file.size);
    var blob = file.slice(offset, end);
    offset = end;

    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var raw = new Uint8Array(e.target.result);

        // Prepend any bytes carried over from the previous chunk
        var combined;
        if (carryover.length > 0) {
          combined = new Uint8Array(carryover.length + raw.length);
          combined.set(carryover);
          combined.set(raw, carryover.length);
        } else {
          combined = raw;
        }

        // Find the last complete character boundary so we don't split
        // a multi‑byte character across chunks
        var split = findLastCompleteBoundary(combined, enc);

        var text = decoder.decode(combined.subarray(0, split));
        carryover = combined.subarray(split);

        if (firstChunk && text) {
          firstChunk = false;
          if (text.charCodeAt(0) === 0xfeff) {
            text = text.slice(1);
          }
        }

        if (text) onChunk(text, offset, file.size);
        setTimeout(next, 0);
      } catch (err) {
        onError(err);
      }
    };
    reader.onerror = function (e) { onError(e); };
    reader.readAsArrayBuffer(blob);
  }

  next();
}

/**
 * UTF‑8 fast path: TextDecoder with {stream:true} handles chunk boundaries
 * and BOM stripping natively and reliably.
 */
function readFileInChunksUTF8(file, onChunk, onDone, onError) {
  var offset = 0;
  var decoder = createTextDecoder("utf-8");
  var firstChunk = true;

  function next() {
    if (offset >= file.size) {
      onDone();
      return;
    }

    var end = Math.min(offset + FILE_CHUNK_SIZE, file.size);
    var blob = file.slice(offset, end);
    offset = end;

    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var text = decoder.decode(e.target.result, { stream: true });

        if (firstChunk) {
          firstChunk = false;
          if (text.charCodeAt(0) === 0xfeff) {
            text = text.slice(1);
          }
        }

        onChunk(text, offset, file.size);
        setTimeout(next, 0);
      } catch (err) {
        onError(err);
      }
    };
    reader.onerror = function (e) { onError(e); };
    reader.readAsArrayBuffer(blob);
  }

  next();
}

/**
 * Create a TextDecoder, falling back to UTF‑8 if the requested encoding
 * is unsupported by the browser.
 */
function createTextDecoder(encoding) {
  try {
    return new TextDecoder(encoding);
  } catch (e) {
    return new TextDecoder("utf-8");
  }
}

/**
 * Find the last byte index in `buffer` that is a safe character boundary
 * for the given encoding. Bytes after this index form an incomplete
 * multi‑byte character and should be carried over to the next chunk.
 *
 * GBK / GB2312 / GB18030:
 *   - 1‑byte: 0x00–0x7F  (ASCII)
 *   - 2‑byte leader: 0x81–0xFE → needs a second byte in [0x40–0xFE]\{0x7F}
 *
 * Big5:
 *   - 1‑byte: 0x00–0x7F
 *   - 2‑byte leader: 0xA1–0xF9 → second byte in [0x40–0x7E] or [0xA1–0xFE]
 */
function findLastCompleteBoundary(buffer, encoding) {
  var len = buffer.length;

  // Single‑byte encodings and ASCII‑subset: everything is complete
  if (encoding === "utf-8" || encoding === "utf8" ||
      encoding === "ascii" || encoding === "latin1" || encoding === "iso-8859-1") {
    return len;
  }

  // Check if it's a GBK-family encoding
  var isGBK = (encoding === "gbk" || encoding === "gb2312" || encoding === "gb18030");
  var isBig5 = (encoding === "big5" || encoding === "big5-hkscs");

  if (!isGBK && !isBig5) {
    // Unknown multi‑byte encoding — conservatively keep the last byte
    return Math.max(0, len - 1);
  }

  // Scan backwards to find the last complete character
  for (var i = len - 1; i >= 0; i--) {
    var b = buffer[i];

    // ASCII byte — always a complete 1‑byte character
    if (b <= 0x7F) {
      return i + 1;
    }

    // Check if this byte is a multi‑byte leader
    var isLeader = false;
    if (isGBK) {
      isLeader = (b >= 0x81 && b <= 0xFE);
    } else if (isBig5) {
      isLeader = (b >= 0xA1 && b <= 0xF9);
    }

    if (isLeader) {
      // Need a second byte after this position
      if (i + 1 < len) {
        return i + 2; // complete 2‑byte char
      } else {
        return i; // incomplete — carry over this byte
      }
    }
    // b is in the second‑byte range → skip, continue backwards
  }

  return 0;
}

// ── Incremental CSV parser (stateful, chunk‑boundary‑safe) ───────────

/**
 * Create a stateful CSV parser that accepts text chunks and maintains
 * state across chunks.  Because the chunk boundary can fall inside a
 * quoted field or a \r\n pair, the parser carries over incomplete state.
 *
 * @param {string} delimiter - field delimiter
 * @param {number} batchSize  - rows to accumulate before flushing
 * @param {function} onBatch  - called (rows, batchIndex) when batch is full
 * @param {boolean} scanOnly  - if true, only count rows / cols (no arrays)
 */
function createParser(delimiter, batchSize, onBatch, scanOnly) {
  var sep = delimiter || ",";

  // ── mutable state persisted across chunks ──
  var batch = scanOnly ? null : [];
  var row = scanOnly ? null : [];
  var field = "";
  var inQuotes = false;
  var batchIndex = 0;
  var totalRows = 0;
  var maxCols = 0;
  var currentCols = 0;
  var firstRowComplete = false;
  // carry‐over for \r\n straddling a chunk boundary
  var pendingCR = false;
  // for scanOnly: true if there is non‑separator content after the last newline
  var hasContentAfterNewline = false;

  function flushBatch() {
    if (!scanOnly && batch && batch.length > 0) {
      // pad rows to maxCols so Excel gets a uniform 2‑D array
      for (var b = 0; b < batch.length; b++) {
        var r = batch[b];
        while (r.length < maxCols) {
          r.push("");
        }
      }
      onBatch(batch, batchIndex);
      batchIndex++;
      batch = [];
    }
  }

  function finishRow() {
    if (scanOnly) {
      if (!firstRowComplete) {
        maxCols = currentCols;
        firstRowComplete = true;
      }
      currentCols = 0;
      hasContentAfterNewline = false;
    } else {
      row.push(field);
      if (row.length > maxCols) {
        maxCols = row.length;
      }
      batch.push(row);
      row = [];
    }
    field = "";
    totalRows++;
    if (!scanOnly && batch.length >= batchSize) {
      flushBatch();
    }
  }

  function feed(text) {
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];

      // ── handle \r\n that was split across chunks ──
      if (pendingCR) {
        pendingCR = false;
        if (ch === "\n") {
          continue; // already handled the \r in previous chunk
        }
        // standalone \r — finish the row, then process current char
        finishRow();
      }

      if (inQuotes) {
        if (ch === '"') {
          var nextCh = text[i + 1];
          if (nextCh === '"') {
            if (!scanOnly) field += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          if (!scanOnly) field += ch;
          else hasContentAfterNewline = true;
        }
      } else {
        if (ch === '"' && field === "") {
          inQuotes = true;
        } else if (ch === sep) {
          if (scanOnly) {
            currentCols++;
          } else {
            row.push(field);
          }
          field = "";
        } else if (ch === "\n") {
          if (scanOnly) currentCols++;
          finishRow();
        } else if (ch === "\r") {
          if (text[i + 1] === "\n") {
            i++; // consume \n
          } else if (i === text.length - 1) {
            // \r at the very end of a chunk — could be \r\n split
            pendingCR = true;
            if (scanOnly) currentCols++;
            finishRow();
            continue;
          }
          if (!pendingCR) {
            if (scanOnly) currentCols++;
            finishRow();
          }
        } else {
          if (!scanOnly) field += ch;
          else hasContentAfterNewline = true;
        }
      }
    }
  }

  function finish() {
    // handle pending \r (unlikely but safe)
    if (pendingCR) {
      pendingCR = false;
    }

    // final cell
    if (scanOnly) {
      if (!firstRowComplete) {
        // Only count a row if there was any input at all
        if (totalRows === 0 && !hasContentAfterNewline && currentCols === 0) {
          return { rowCount: 0, colCount: 0 };
        }
        currentCols++;
        maxCols = currentCols;
        totalRows++;
      }
      // If firstRowComplete: rows before the last \n were already counted.
      // Count the final row only if there is content after the last newline.
      else if (hasContentAfterNewline) {
        totalRows++;
      }
    } else {
      row.push(field);
      // Check whether this final row has any real content
      var hasContent = false;
      for (var c = 0; c < row.length; c++) {
        if (row[c] !== "") {
          hasContent = true;
          break;
        }
      }
      if (hasContent) {
        if (row.length > maxCols) {
          maxCols = row.length;
        }
        batch.push(row);
        totalRows++;
      }
      flushBatch();
    }

    field = "";
    return { rowCount: totalRows, colCount: maxCols };
  }

  return { feed: feed, finish: finish, getCounts: function () { return { rowCount: totalRows, colCount: maxCols }; } };
}

// ── Public streaming API ──────────────────────────────────────────────

/**
 * Scan a CSV file to count rows and columns WITHOUT building any arrays.
 * Reads the file in chunks; peak memory ≈ one 1-MB chunk + counters.
 *
 * @param {File} file       - Browser File object
 * @param {string} delimiter
 * @param {string} encoding  - Text encoding (default "utf-8"; use "gbk" for GBK)
 * @param {function} onDone  - called ({rowCount, colCount}) on success
 * @param {function} onError - called (err) on failure
 */
function scanCSVFile(file, delimiter, encoding, onDone, onError) {
  if (!file || file.size === 0) {
    onDone({ rowCount: 0, colCount: 0 });
    return;
  }

  var parser = createParser(delimiter, 0, null, /* scanOnly */ true);

  readFileInChunks(
    file,
    encoding,
    function (text, bytesRead, totalBytes) {
      parser.feed(text);
    },
    function () {
      var stats = parser.finish();
      onDone(stats);
    },
    onError || function () {}
  );
}

/**
 * Parse a CSV file in batches, writing each batch via onBatch.
 * Reads the file in chunks; peak memory ≈ one 1-MB chunk + one batch of rows.
 *
 * @param {File} file       - Browser File object
 * @param {string} delimiter
 * @param {number} batchSize- rows to accumulate before onBatch
 * @param {string} encoding  - Text encoding (default "utf-8"; use "gbk" for GBK)
 * @param {function} onBatch - called (rows, batchIndex) for each batch
 * @param {function} onDone  - called ({rowCount, colCount}) on success
 * @param {function} onError - called (err) on failure
 */
function parseCSVFile(file, delimiter, batchSize, encoding, onBatch, onDone, onError) {
  if (!file || file.size === 0) {
    onDone({ rowCount: 0, colCount: 0 });
    return;
  }

  var parser = createParser(delimiter, batchSize, onBatch, /* scanOnly */ false);

  readFileInChunks(
    file,
    encoding,
    function (text) {
      parser.feed(text);
    },
    function () {
      var stats = parser.finish();
      onDone(stats);
    },
    onError || function () {}
  );
}

module.exports = {
  parseCSV: parseCSV,
  scanCSVFile: scanCSVFile,
  parseCSVFile: parseCSVFile,
  createParser: createParser, // exported for testing
};
