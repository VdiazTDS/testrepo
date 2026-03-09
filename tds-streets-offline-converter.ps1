param(
  [string]$ZipPath,
  [string]$OutputGeoJson
)

$ErrorActionPreference = "Stop"

function Write-Step([string]$Message) {
  Write-Host "[TDS Streets Converter] $Message"
}

function Show-OpenFileDialog([string]$Title, [string]$Filter) {
  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  $dialog = New-Object System.Windows.Forms.OpenFileDialog
  $dialog.Title = $Title
  $dialog.Filter = $Filter
  $dialog.Multiselect = $false
  $dialog.CheckFileExists = $true
  $result = $dialog.ShowDialog()
  if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    return $null
  }
  return $dialog.FileName
}

function Show-SaveFileDialog([string]$Title, [string]$Filter, [string]$DefaultPath) {
  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  $dialog = New-Object System.Windows.Forms.SaveFileDialog
  $dialog.Title = $Title
  $dialog.Filter = $Filter
  $dialog.FileName = [System.IO.Path]::GetFileName($DefaultPath)
  $dialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($DefaultPath)
  $result = $dialog.ShowDialog()
  if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    return $null
  }
  return $dialog.FileName
}

function Test-PythonInvocation([string]$Command, [string[]]$PrefixArgs) {
  try {
    $args = @($PrefixArgs + @("-c", "import sys; print(sys.version)"))
    & $Command @args *> $null
    return ($LASTEXITCODE -eq 0)
  } catch {
    return $false
  }
}

function Get-PythonInvocation {
  $candidates = @(
    @{ Command = "python"; PrefixArgs = @() },
    @{ Command = "python3"; PrefixArgs = @() },
    @{ Command = "py"; PrefixArgs = @("-3") }
  )

  foreach ($candidate in $candidates) {
    $cmdName = [string]$candidate.Command
    $prefix = @($candidate.PrefixArgs)
    $found = Get-Command $cmdName -ErrorAction SilentlyContinue
    if (-not $found) { continue }
    if (Test-PythonInvocation -Command $cmdName -PrefixArgs $prefix) {
      return @{
        Command = $cmdName
        PrefixArgs = $prefix
      }
    }
  }
  return $null
}

function Ensure-PyShpInstalled($PythonInvocation) {
  $cmd = [string]$PythonInvocation.Command
  $prefix = @($PythonInvocation.PrefixArgs)
  $displayCmd = "$cmd $($prefix -join ' ')".Trim()

  & $cmd @($prefix + @("-c", "import shapefile")) 2>$null
  if ($LASTEXITCODE -eq 0) { return }

  Write-Step "Installing Python package: pyshp"
  & $cmd @($prefix + @("-m", "pip", "install", "--user", "pyshp"))
  if ($LASTEXITCODE -ne 0) {
    throw "Unable to install pyshp. Run: $displayCmd -m pip install --user pyshp"
  }
}

function Get-NodeCommand {
  $node = Get-Command node -ErrorAction SilentlyContinue
  if (-not $node) { return $null }
  try {
    & node -e "process.exit(0)" *> $null
    if ($LASTEXITCODE -eq 0) { return "node" }
  } catch {}
  return $null
}

function Get-ShapefileScore([string]$Name) {
  $name = $Name.ToLowerInvariant()
  $score = 0
  if ($name -eq "gis_osm_roads_free_1.shp") { $score += 1200 }
  if ($name -eq "gis_osm_highways_free_1.shp") { $score += 900 }
  if ($name -like "*roads_free_1.shp") { $score += 700 }
  if ($name -like "*roads*.shp") { $score += 300 }
  if ($name -like "*highway*.shp") { $score += 220 }
  if ($name -like "*street*.shp") { $score += 160 }
  if ($name -like "*rail*.shp") { $score -= 120 }
  if ($name -like "*water*.shp") { $score -= 120 }
  if ($name -like "*building*.shp") { $score -= 120 }
  return $score
}

function New-WorkingTempDirectory([string]$OutputPath) {
  $candidates = @()
  if ($OutputPath) {
    $outDir = Split-Path -Parent $OutputPath
    if ($outDir) {
      $candidates += $outDir
    }
  }
  if ($env:TEMP) {
    $candidates += $env:TEMP
  }

  foreach ($candidate in $candidates) {
    if (-not $candidate) { continue }
    try {
      if (-not (Test-Path -LiteralPath $candidate)) { continue }
      $tmp = Join-Path $candidate ("tds_streets_convert_" + [Guid]::NewGuid().ToString("N"))
      New-Item -Path $tmp -ItemType Directory -ErrorAction Stop | Out-Null
      return $tmp
    } catch {}
  }

  throw "Could not create a temporary folder near output path or TEMP."
}

function Extract-RoadShapefileFromZip([string]$ZipPathValue, [string]$DestinationRoot) {
  Add-Type -AssemblyName System.IO.Compression | Out-Null
  Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

  $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipPathValue)
  try {
    $shpEntries = @(
      $archive.Entries | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.Name) -and $_.Name.ToLowerInvariant().EndsWith(".shp")
      }
    )
    if (-not $shpEntries -or $shpEntries.Count -eq 0) {
      throw "No .shp files found in ZIP."
    }

    $ranked = $shpEntries | ForEach-Object {
      [PSCustomObject]@{
        Entry = $_
        Score = Get-ShapefileScore -Name $_.Name
      }
    } | Sort-Object -Property Score -Descending

    $best = $ranked | Select-Object -First 1
    if (-not $best) {
      throw "Unable to choose a roads shapefile from ZIP."
    }

    $bestEntry = $best.Entry
    $targetBase = [System.IO.Path]::GetFileNameWithoutExtension($bestEntry.Name)
    $targetDir = [System.IO.Path]::GetDirectoryName($bestEntry.FullName)
    if (-not $targetDir) { $targetDir = "" }
    $keepExt = @(".shp", ".shx", ".dbf", ".prj", ".cpg", ".qix", ".fix")

    $toExtract = @(
      $archive.Entries | Where-Object {
        if ([string]::IsNullOrWhiteSpace($_.Name)) { return $false }
        $entryDir = [System.IO.Path]::GetDirectoryName($_.FullName)
        if (-not $entryDir) { $entryDir = "" }
        if ($entryDir -ne $targetDir) { return $false }
        $entryBase = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        if ($entryBase -ne $targetBase) { return $false }
        $entryExt = [System.IO.Path]::GetExtension($_.Name).ToLowerInvariant()
        return ($keepExt -contains $entryExt)
      }
    )

    if (-not $toExtract -or $toExtract.Count -eq 0) {
      throw "Required shapefile sidecar files were not found in ZIP."
    }

    foreach ($entry in $toExtract) {
      $destPath = Join-Path $DestinationRoot $entry.FullName
      $destDir = Split-Path -Parent $destPath
      if ($destDir -and -not (Test-Path -LiteralPath $destDir)) {
        New-Item -Path $destDir -ItemType Directory -Force | Out-Null
      }
      [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $destPath, $true)
    }

    $shpOut = Join-Path $DestinationRoot $bestEntry.FullName
    if (-not (Test-Path -LiteralPath $shpOut)) {
      throw "Shapefile extraction failed: $shpOut"
    }

    Write-Step ("Selected shapefile: " + $bestEntry.Name)
    return $shpOut
  } finally {
    if ($archive) { $archive.Dispose() }
  }
}

function Find-RoadShapefile([string]$RootDir) {
  $all = Get-ChildItem -Path $RootDir -Recurse -File -Filter *.shp
  if (-not $all -or $all.Count -eq 0) {
    throw "No .shp files found after extracting ZIP."
  }

  $scored = $all | ForEach-Object {
    [PSCustomObject]@{
      Path = $_.FullName
      Score = Get-ShapefileScore -Name $_.Name
      Name = $_.Name
    }
  } | Sort-Object -Property Score -Descending

  $best = $scored | Select-Object -First 1
  if (-not $best) {
    throw "Unable to choose a roads shapefile."
  }
  Write-Step ("Selected shapefile: " + $best.Name)
  return $best.Path
}

function Build-PythonConverterScript {
@'
import json
import os
import sys
import shapefile

if len(sys.argv) < 3:
    print("Usage: converter.py <input_shp> <output_geojson>")
    sys.exit(2)

input_shp = sys.argv[1]
output_geojson = sys.argv[2]
COORD_PRECISION = 6

def round_coord(v):
    return round(float(v), COORD_PRECISION)

def clean_value(v):
    if v is None:
        return ""
    if isinstance(v, bytes):
        try:
            v = v.decode("utf-8", errors="ignore")
        except Exception:
            v = str(v)
    if isinstance(v, float):
        if not (v == v):  # NaN
            return ""
        return round(v, COORD_PRECISION)
    if isinstance(v, (int, bool)):
        return v
    return str(v).strip()

def pick_value(props, *keys, fallback="Unknown"):
    for key in keys:
        if key in props and props[key] not in (None, ""):
            return clean_value(props[key])
    return fallback

reader = shapefile.Reader(input_shp, encoding="utf-8", encodingErrors="ignore")
total = reader.numRecords
field_names = [f[0] for f in reader.fields[1:]]

os.makedirs(os.path.dirname(output_geojson) or ".", exist_ok=True)

with open(output_geojson, "w", encoding="utf-8") as f:
    f.write('{"type":"FeatureCollection","features":[')
    first = True
    kept = 0
    for idx, sr in enumerate(reader.iterShapeRecords(), start=1):
        shape = sr.shape
        st = int(getattr(shape, "shapeType", 0))
        if st not in (3, 13, 23):
            continue
        pts = list(shape.points or [])
        if len(pts) < 2:
            continue
        parts = list(shape.parts or [0])
        if not parts:
            parts = [0]
        parts.append(len(pts))
        lines = []
        for i in range(len(parts) - 1):
            seg = pts[parts[i]:parts[i + 1]]
            if len(seg) < 2:
                continue
            coords = []
            for x, y in seg:
                coords.append([round_coord(x), round_coord(y)])
            if len(coords) >= 2:
                lines.append(coords)
        if not lines:
            continue

        geom = {"type": "LineString", "coordinates": lines[0]} if len(lines) == 1 else {"type": "MultiLineString", "coordinates": lines}
        rec_values = list(sr.record) if sr.record is not None else []
        rec_props = dict(zip(field_names, rec_values))
        normalized_props = {
            "name": pick_value(rec_props, "name", "NAME", "osm_name", "OSM_NAME", fallback="Unknown"),
            "highway": pick_value(rec_props, "highway", "HIGHWAY", "fclass", "FCLASS", "type", "TYPE", fallback="Unknown"),
            "ref": pick_value(rec_props, "ref", "REF", "ref_name", "REF_NAME", fallback="Unknown"),
            "maxspeed": pick_value(rec_props, "maxspeed", "MAXSPEED", "max_speed", "MAX_SPEED", fallback="Unknown"),
            "lanes": pick_value(rec_props, "lanes", "LANES", "num_lanes", "NUM_LANES", fallback="Unknown"),
            "surface": pick_value(rec_props, "surface", "SURFACE", "surf_type", "SURF_TYPE", fallback="Unknown"),
            "oneway": pick_value(rec_props, "oneway", "ONEWAY", "one_way", "ONE_WAY", fallback="Unknown")
        }

        feature = {
            "type": "Feature",
            "id": idx,
            "properties": normalized_props,
            "geometry": geom
        }
        if not first:
            f.write(",")
        first = False
        json.dump(feature, f, separators=(",", ":"))
        kept += 1

        if kept % 25000 == 0:
            pct = (idx / total * 100.0) if total else 0.0
            print(f"Progress: {pct:.1f}% ({kept} roads written)")

    f.write("]}")

print(f"Done. Wrote {kept} street features to {output_geojson}")
'@
}

function Build-NodeConverterScript {
@'
const fs = require("fs");
const path = require("path");

if (process.argv.length < 4) {
  console.error("Usage: node converter.js <input_shp> <output_geojson>");
  process.exit(2);
}

const inputShp = process.argv[2];
const outputGeoJson = process.argv[3];
const inputDbf = inputShp.replace(/\.shp$/i, ".dbf");

function fail(msg) {
  console.error(msg);
  process.exit(1);
}

if (!fs.existsSync(inputShp)) {
  fail(`Shapefile not found: ${inputShp}`);
}

const outDir = path.dirname(outputGeoJson);
if (outDir && outDir !== ".") {
  fs.mkdirSync(outDir, { recursive: true });
}

const shpFd = fs.openSync(inputShp, "r");
const shpSize = fs.fstatSync(shpFd).size;
if (shpSize < 100) {
  fs.closeSync(shpFd);
  fail("Invalid shapefile (too small).");
}

const outFd = fs.openSync(outputGeoJson, "w");
const COORD_PRECISION = 1e6;
let pos = 100;
let recordId = 0;
let written = 0;
let first = true;
let buffer = '{"type":"FeatureCollection","features":[';
const FLUSH_SIZE = 4 * 1024 * 1024;
let dbfFd = null;
let dbfMeta = null;

function flush(force) {
  if (!force && buffer.length < FLUSH_SIZE) return;
  if (!buffer.length) return;
  fs.writeSync(outFd, buffer, null, "utf8");
  buffer = "";
}

function readExact(length, position) {
  const chunk = Buffer.alloc(length);
  const bytesRead = fs.readSync(shpFd, chunk, 0, length, position);
  if (bytesRead !== length) {
    throw new Error("Unexpected end of shapefile while reading records.");
  }
  return chunk;
}

function parseDbfHeader(fd) {
  const header = Buffer.alloc(32);
  const read = fs.readSync(fd, header, 0, 32, 0);
  if (read !== 32) {
    throw new Error("DBF header is incomplete.");
  }
  const recordCount = header.readUInt32LE(4);
  const headerLength = header.readUInt16LE(8);
  const recordLength = header.readUInt16LE(10);
  if (headerLength < 33 || recordLength < 2) {
    throw new Error("DBF header is invalid.");
  }
  const fields = [];
  let offset = 32;
  while (offset + 32 <= headerLength) {
    const desc = Buffer.alloc(32);
    const n = fs.readSync(fd, desc, 0, 32, offset);
    if (n !== 32) break;
    if (desc[0] === 0x0d) break;
    const zero = desc.indexOf(0x00, 0);
    const rawName = desc.slice(0, zero >= 0 ? zero : 11).toString("ascii").trim();
    const fieldName = rawName || `field_${fields.length + 1}`;
    const fieldType = String.fromCharCode(desc[11] || 67);
    const fieldLength = desc[16] || 0;
    fields.push({ name: fieldName, type: fieldType, length: fieldLength });
    offset += 32;
  }
  return { recordCount, headerLength, recordLength, fields };
}

function decodeDbfValue(rawBuf, fieldType) {
  const text = rawBuf.toString("utf8").replace(/\u0000+/g, "").trim();
  if (!text) return "";
  if (fieldType === "N" || fieldType === "F") {
    const num = Number(text);
    if (Number.isFinite(num)) return String(num);
  }
  return text;
}

function readDbfRecord(index) {
  if (!dbfFd || !dbfMeta) return {};
  if (!Number.isFinite(index) || index < 0 || index >= dbfMeta.recordCount) return {};
  const rowPos = dbfMeta.headerLength + (index * dbfMeta.recordLength);
  const row = Buffer.alloc(dbfMeta.recordLength);
  const read = fs.readSync(dbfFd, row, 0, row.length, rowPos);
  if (read !== row.length) return {};
  if (row[0] === 0x2a) return {}; // deleted record

  const props = {};
  let cursor = 1;
  dbfMeta.fields.forEach(field => {
    const len = Number(field.length || 0);
    if (len <= 0 || (cursor + len) > row.length) {
      cursor += Math.max(0, len);
      return;
    }
    const value = decodeDbfValue(row.slice(cursor, cursor + len), field.type);
    props[field.name] = value;
    props[field.name.toLowerCase()] = value;
    cursor += len;
  });
  return props;
}

function pickValue(props, keys) {
  for (let i = 0; i < keys.length; i += 1) {
    const key = keys[i];
    const value = props[key];
    if (value === undefined || value === null) continue;
    const text = String(value).trim();
    if (text) return text;
  }
  return "Unknown";
}

if (fs.existsSync(inputDbf)) {
  try {
    dbfFd = fs.openSync(inputDbf, "r");
    dbfMeta = parseDbfHeader(dbfFd);
    console.log(`DBF attributes loaded (${dbfMeta.fields.length} fields).`);
  } catch (err) {
    if (dbfFd) {
      try { fs.closeSync(dbfFd); } catch (_) {}
      dbfFd = null;
    }
    dbfMeta = null;
    console.warn(`DBF parse failed: ${err.message}. Continuing with Unknown attributes.`);
  }
} else {
  console.warn("DBF file not found. Attributes will be Unknown.");
}

try {
  while (pos + 8 <= shpSize) {
    const recHeader = readExact(8, pos);
    const contentLengthWords = recHeader.readUInt32BE(4);
    const contentLengthBytes = contentLengthWords * 2;
    pos += 8;

    if (contentLengthBytes <= 0) {
      continue;
    }
    if (pos + contentLengthBytes > shpSize) {
      console.warn("Shapefile ended with a truncated record. Stopping parse.");
      break;
    }

    const rec = readExact(contentLengthBytes, pos);
    pos += contentLengthBytes;
    recordId += 1;

    if (rec.length < 44) continue;
    const shapeType = rec.readInt32LE(0);
    if (shapeType !== 3 && shapeType !== 13 && shapeType !== 23) continue;

    const numParts = rec.readInt32LE(36);
    const numPoints = rec.readInt32LE(40);
    if (numParts <= 0 || numPoints <= 1) continue;

    const partsOffset = 44;
    const pointsOffset = partsOffset + (numParts * 4);
    if (pointsOffset + (numPoints * 16) > rec.length) continue;

    const parts = [];
    let valid = true;
    for (let i = 0; i < numParts; i += 1) {
      const idx = rec.readInt32LE(partsOffset + (i * 4));
      if (idx < 0 || idx >= numPoints) {
        valid = false;
        break;
      }
      parts.push(idx);
    }
    if (!valid) continue;
    parts.push(numPoints);

    const lines = [];
    for (let i = 0; i < parts.length - 1; i += 1) {
      const start = parts[i];
      const end = parts[i + 1];
      if (end - start < 2) continue;

      const coords = [];
      for (let p = start; p < end; p += 1) {
        const offset = pointsOffset + (p * 16);
        const x = Math.round(rec.readDoubleLE(offset) * COORD_PRECISION) / COORD_PRECISION;
        const y = Math.round(rec.readDoubleLE(offset + 8) * COORD_PRECISION) / COORD_PRECISION;
        if (!Number.isFinite(x) || !Number.isFinite(y)) continue;
        coords.push([x, y]);
      }
      if (coords.length >= 2) {
        lines.push(coords);
      }
    }
    if (!lines.length) continue;

    const geometry = lines.length === 1
      ? { type: "LineString", coordinates: lines[0] }
      : { type: "MultiLineString", coordinates: lines };

    const dbfProps = readDbfRecord(recordId - 1);
    const feature = {
      type: "Feature",
      id: recordId,
      properties: {
        name: pickValue(dbfProps, ["name", "NAME", "osm_name", "OSM_NAME"]),
        highway: pickValue(dbfProps, ["highway", "HIGHWAY", "fclass", "FCLASS", "type", "TYPE", "road_class", "ROAD_CLASS"]),
        ref: pickValue(dbfProps, ["ref", "REF", "ref_name", "REF_NAME"]),
        maxspeed: pickValue(dbfProps, ["maxspeed", "MAXSPEED", "max_speed", "MAX_SPEED"]),
        lanes: pickValue(dbfProps, ["lanes", "LANES", "num_lanes", "NUM_LANES"]),
        surface: pickValue(dbfProps, ["surface", "SURFACE", "surf_type", "SURF_TYPE"]),
        oneway: pickValue(dbfProps, ["oneway", "ONEWAY", "one_way", "ONE_WAY"])
      },
      geometry
    };

    if (!first) buffer += ",";
    first = false;
    buffer += JSON.stringify(feature);
    written += 1;

    flush(false);

    if (recordId % 25000 === 0) {
      const pct = Math.min(100, (pos / shpSize) * 100);
      console.log(`Progress: ${pct.toFixed(1)}% (${written} roads written)`);
    }
  }

  buffer += "]}";
  flush(true);
} catch (err) {
  try { fs.closeSync(shpFd); } catch (_) {}
  try { fs.closeSync(outFd); } catch (_) {}
  if (dbfFd) {
    try { fs.closeSync(dbfFd); } catch (_) {}
  }
  throw err;
}

fs.closeSync(shpFd);
fs.closeSync(outFd);
if (dbfFd) {
  fs.closeSync(dbfFd);
}
console.log(`Done. Wrote ${written} street features to ${outputGeoJson}`);
'@
}

function Format-GeoJsonNumber([double]$Value) {
  if ([double]::IsNaN($Value) -or [double]::IsInfinity($Value)) {
    return $null
  }
  $rounded = [Math]::Round($Value, 6, [MidpointRounding]::AwayFromZero)
  return $rounded.ToString("0.######", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Convert-ShapefileToGeoJsonPowerShell([string]$ShpPath, [string]$OutputGeoJson) {
  if (-not (Test-Path -LiteralPath $ShpPath)) {
    throw "Shapefile not found: $ShpPath"
  }

  $outDir = Split-Path -Parent $OutputGeoJson
  if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -Path $outDir -ItemType Directory -Force | Out-Null
  }

  $inStream = $null
  $reader = $null
  $outStream = $null
  $writer = $null
  $dbfStream = $null
  $dbfReader = $null
  $dbfMeta = $null
  $dbfFieldDefs = @()
  $dbfPath = [System.IO.Path]::ChangeExtension($ShpPath, ".dbf")

  function Escape-JsonStringValue([string]$Text) {
    $s = [string]$Text
    $s = $s -replace '\\', '\\\\'
    $s = $s -replace '"', '\"'
    $s = $s -replace "`r", '\r'
    $s = $s -replace "`n", '\n'
    $s = $s -replace "`t", '\t'
    return '"' + $s + '"'
  }

  function Parse-DbfHeader([System.IO.BinaryReader]$DbfReader) {
    $header = $DbfReader.ReadBytes(32)
    if ($header.Length -lt 32) {
      throw "DBF header is incomplete."
    }
    $recordCount = [System.BitConverter]::ToUInt32($header, 4)
    $headerLength = [System.BitConverter]::ToUInt16($header, 8)
    $recordLength = [System.BitConverter]::ToUInt16($header, 10)
    if ($headerLength -lt 33 -or $recordLength -lt 2) {
      throw "DBF header is invalid."
    }

    $fields = New-Object System.Collections.Generic.List[hashtable]
    $offset = 32
    while (($offset + 32) -le $headerLength) {
      $desc = $DbfReader.ReadBytes(32)
      if ($desc.Length -lt 32) { break }
      $offset += 32
      if ($desc[0] -eq 0x0D) { break }

      $zeroIdx = [System.Array]::IndexOf($desc, [byte]0, 0, 11)
      if ($zeroIdx -lt 0) { $zeroIdx = 11 }
      $nameLen = [Math]::Min([Math]::Max($zeroIdx, 0), 11)
      $nameBytes = if ($nameLen -gt 0) { $desc[0..($nameLen - 1)] } else { @() }
      $fieldName = ([System.Text.Encoding]::ASCII.GetString($nameBytes)).Trim([char]0).Trim()
      if ([string]::IsNullOrWhiteSpace($fieldName)) {
        $fieldName = "field_$($fields.Count + 1)"
      }

      $fieldType = [char]$desc[11]
      $fieldLength = [int]$desc[16]
      $fields.Add(@{
        name = $fieldName
        type = $fieldType
        length = $fieldLength
      })
    }

    return @{
      recordCount = [int]$recordCount
      headerLength = [int]$headerLength
      recordLength = [int]$recordLength
      fields = $fields
    }
  }

  function Decode-DbfFieldValue([byte[]]$Raw, [string]$FieldType) {
    if ($null -eq $Raw -or $Raw.Length -eq 0) { return "" }
    $text = [System.Text.Encoding]::UTF8.GetString($Raw)
    if ($null -eq $text) { return "" }
    $text = $text.Replace([char]0, "").Trim()
    if (-not $text) { return "" }
    if ($FieldType -eq "N" -or $FieldType -eq "F") {
      $num = 0.0
      if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
        return $num.ToString("0.################", [System.Globalization.CultureInfo]::InvariantCulture)
      }
    }
    return $text
  }

  function Read-DbfRecord([int]$RowIndex) {
    if (-not $dbfReader -or -not $dbfMeta) { return @{} }
    if ($RowIndex -lt 0 -or $RowIndex -ge [int]$dbfMeta.recordCount) { return @{} }

    $rowPos = [int64]$dbfMeta.headerLength + ([int64]$RowIndex * [int64]$dbfMeta.recordLength)
    if ($rowPos -lt 0 -or $rowPos -ge [int64]$dbfReader.BaseStream.Length) { return @{} }
    $dbfReader.BaseStream.Position = $rowPos
    $row = $dbfReader.ReadBytes([int]$dbfMeta.recordLength)
    if ($row.Length -lt [int]$dbfMeta.recordLength) { return @{} }
    if ($row[0] -eq 0x2A) { return @{} } # deleted record

    $props = @{}
    $cursor = 1
    foreach ($field in $dbfFieldDefs) {
      $len = [int]$field.length
      if ($len -le 0 -or ($cursor + $len) -gt $row.Length) {
        $cursor += [Math]::Max(0, $len)
        continue
      }
      $value = Decode-DbfFieldValue -Raw $row[$cursor..($cursor + $len - 1)] -FieldType ([string]$field.type)
      $props[$field.name] = $value
      $props[([string]$field.name).ToLowerInvariant()] = $value
      $cursor += $len
    }
    return $props
  }

  function Pick-DbfValue([hashtable]$Props, [string[]]$Keys, [string]$Fallback = "Unknown") {
    foreach ($key in $Keys) {
      if (-not $Props.ContainsKey($key)) { continue }
      $raw = $Props[$key]
      if ($null -eq $raw) { continue }
      $text = [string]$raw
      if ($null -eq $text) { continue }
      $text = $text.Trim()
      if ($text) { return $text }
    }
    return $Fallback
  }

  try {
    $inStream = [System.IO.File]::Open($ShpPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    if ($inStream.Length -lt 100) {
      throw "Invalid shapefile (too small)."
    }
    $reader = New-Object System.IO.BinaryReader($inStream)

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $outStream = [System.IO.File]::Open($OutputGeoJson, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
    $writer = New-Object System.IO.StreamWriter($outStream, $utf8NoBom)

    if (Test-Path -LiteralPath $dbfPath) {
      try {
        $dbfStream = [System.IO.File]::Open($dbfPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $dbfReader = New-Object System.IO.BinaryReader($dbfStream)
        $dbfMeta = Parse-DbfHeader -DbfReader $dbfReader
        $dbfFieldDefs = @($dbfMeta.fields)
        Write-Step ("DBF attributes loaded ({0} fields)." -f $dbfFieldDefs.Count)
      } catch {
        Write-Step ("DBF parse failed: {0}. Continuing with Unknown attributes." -f $_.Exception.Message)
        if ($dbfReader) { $dbfReader.Dispose(); $dbfReader = $null }
        if ($dbfStream) { $dbfStream.Dispose(); $dbfStream = $null }
        $dbfMeta = $null
        $dbfFieldDefs = @()
      }
    } else {
      Write-Step "DBF file not found. Attributes will be Unknown."
    }

    $writer.Write('{"type":"FeatureCollection","features":[')
    $firstFeature = $true
    $recordId = 0
    $written = 0

    $inStream.Position = 100
    while (($inStream.Position + 8) -le $inStream.Length) {
      $recordNumberBytes = $reader.ReadBytes(4)
      if ($recordNumberBytes.Length -lt 4) { break }
      [Array]::Reverse($recordNumberBytes)

      $contentLengthBytesBe = $reader.ReadBytes(4)
      if ($contentLengthBytesBe.Length -lt 4) { break }
      [Array]::Reverse($contentLengthBytesBe)
      $contentLengthWords = [System.BitConverter]::ToInt32($contentLengthBytesBe, 0)
      $contentLength = $contentLengthWords * 2
      if ($contentLength -le 0) { continue }

      if (($inStream.Position + $contentLength) -gt $inStream.Length) {
        Write-Step "Shapefile ended with a truncated record. Stopping parse."
        break
      }

      $record = $reader.ReadBytes($contentLength)
      if ($record.Length -lt $contentLength) { break }
      $recordId += 1

      if ($record.Length -lt 44) { continue }
      $shapeType = [System.BitConverter]::ToInt32($record, 0)
      if ($shapeType -ne 3 -and $shapeType -ne 13 -and $shapeType -ne 23) { continue }

      $numParts = [System.BitConverter]::ToInt32($record, 36)
      $numPoints = [System.BitConverter]::ToInt32($record, 40)
      if ($numParts -le 0 -or $numPoints -le 1) { continue }

      $partsOffset = 44
      $pointsOffset = $partsOffset + ($numParts * 4)
      if (($pointsOffset + ($numPoints * 16)) -gt $record.Length) { continue }

      $parts = New-Object System.Collections.Generic.List[int]
      $isValid = $true
      for ($partIdx = 0; $partIdx -lt $numParts; $partIdx++) {
        $idx = [System.BitConverter]::ToInt32($record, $partsOffset + ($partIdx * 4))
        if ($idx -lt 0 -or $idx -ge $numPoints) {
          $isValid = $false
          break
        }
        $parts.Add($idx)
      }
      if (-not $isValid) { continue }
      $parts.Add($numPoints)

      $lineJsonList = New-Object System.Collections.Generic.List[string]
      for ($partIdx = 0; $partIdx -lt ($parts.Count - 1); $partIdx++) {
        $start = $parts[$partIdx]
        $end = $parts[$partIdx + 1]
        if (($end - $start) -lt 2) { continue }

        $coordSb = New-Object System.Text.StringBuilder
        [void]$coordSb.Append("[")
        $coordCount = 0
        for ($pointIdx = $start; $pointIdx -lt $end; $pointIdx++) {
          $offset = $pointsOffset + ($pointIdx * 16)
          $xRaw = [System.BitConverter]::ToDouble($record, $offset)
          $yRaw = [System.BitConverter]::ToDouble($record, $offset + 8)
          $x = Format-GeoJsonNumber -Value $xRaw
          $y = Format-GeoJsonNumber -Value $yRaw
          if ($null -eq $x -or $null -eq $y) { continue }

          if ($coordCount -gt 0) {
            [void]$coordSb.Append(",")
          }
          [void]$coordSb.Append("[")
          [void]$coordSb.Append($x)
          [void]$coordSb.Append(",")
          [void]$coordSb.Append($y)
          [void]$coordSb.Append("]")
          $coordCount += 1
        }

        if ($coordCount -ge 2) {
          [void]$coordSb.Append("]")
          $lineJsonList.Add($coordSb.ToString())
        }
      }

      if ($lineJsonList.Count -eq 0) { continue }

      if ($lineJsonList.Count -eq 1) {
        $geomType = "LineString"
        $coordsJson = $lineJsonList[0]
      } else {
        $geomType = "MultiLineString"
        $coordsJson = "[" + ($lineJsonList -join ",") + "]"
      }

      if (-not $firstFeature) {
        $writer.Write(",")
      } else {
        $firstFeature = $false
      }

      $dbfProps = Read-DbfRecord -RowIndex ($recordId - 1)
      $nameValue = Pick-DbfValue -Props $dbfProps -Keys @("name", "NAME", "osm_name", "OSM_NAME")
      $highwayValue = Pick-DbfValue -Props $dbfProps -Keys @("highway", "HIGHWAY", "fclass", "FCLASS", "type", "TYPE", "road_class", "ROAD_CLASS")
      $refValue = Pick-DbfValue -Props $dbfProps -Keys @("ref", "REF", "ref_name", "REF_NAME")
      $maxspeedValue = Pick-DbfValue -Props $dbfProps -Keys @("maxspeed", "MAXSPEED", "max_speed", "MAX_SPEED")
      $lanesValue = Pick-DbfValue -Props $dbfProps -Keys @("lanes", "LANES", "num_lanes", "NUM_LANES")
      $surfaceValue = Pick-DbfValue -Props $dbfProps -Keys @("surface", "SURFACE", "surf_type", "SURF_TYPE")
      $onewayValue = Pick-DbfValue -Props $dbfProps -Keys @("oneway", "ONEWAY", "one_way", "ONE_WAY")

      $writer.Write('{"type":"Feature","id":')
      $writer.Write($recordId)
      $writer.Write(',"properties":{"name":')
      $writer.Write((Escape-JsonStringValue -Text $nameValue))
      $writer.Write(',"highway":')
      $writer.Write((Escape-JsonStringValue -Text $highwayValue))
      $writer.Write(',"ref":')
      $writer.Write((Escape-JsonStringValue -Text $refValue))
      $writer.Write(',"maxspeed":')
      $writer.Write((Escape-JsonStringValue -Text $maxspeedValue))
      $writer.Write(',"lanes":')
      $writer.Write((Escape-JsonStringValue -Text $lanesValue))
      $writer.Write(',"surface":')
      $writer.Write((Escape-JsonStringValue -Text $surfaceValue))
      $writer.Write(',"oneway":')
      $writer.Write((Escape-JsonStringValue -Text $onewayValue))
      $writer.Write('},"geometry":{"type":"')
      $writer.Write($geomType)
      $writer.Write('","coordinates":')
      $writer.Write($coordsJson)
      $writer.Write("}}")

      $written += 1
      if (($recordId % 25000) -eq 0) {
        $pct = [math]::Min(100, ($inStream.Position / $inStream.Length) * 100)
        Write-Step ("Progress: {0:N1}% ({1} roads written)" -f $pct, $written)
      }
    }

    $writer.Write("]}")
    $writer.Flush()
    Write-Step "Built-in parser wrote $written street features."
  } finally {
    if ($dbfReader) { $dbfReader.Dispose() }
    if ($dbfStream) { $dbfStream.Dispose() }
    if ($writer) { $writer.Dispose() }
    if ($outStream) { $outStream.Dispose() }
    if ($reader) { $reader.Dispose() }
    if ($inStream) { $inStream.Dispose() }
  }
}

function Test-GeoJsonFileLooksComplete([string]$GeoJsonPath) {
  if (-not (Test-Path -LiteralPath $GeoJsonPath)) {
    return $false
  }

  $file = Get-Item -LiteralPath $GeoJsonPath
  if (-not $file -or $file.Length -lt 64) {
    return $false
  }

  $stream = $null
  try {
    $stream = [System.IO.File]::Open($GeoJsonPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $headLength = [Math]::Min(160, [int]$stream.Length)
    $headBytes = New-Object byte[] $headLength
    [void]$stream.Read($headBytes, 0, $headLength)
    $headText = [System.Text.Encoding]::UTF8.GetString($headBytes)
    $headNormalized = ($headText -replace "\s+", "").ToLowerInvariant()
    if (-not $headNormalized.Contains('"type":"featurecollection"')) {
      return $false
    }
    if (-not $headNormalized.Contains('"features":[')) {
      return $false
    }

    $tailLength = [Math]::Min(4096, [int]$stream.Length)
    $stream.Position = $stream.Length - $tailLength
    $tailBytes = New-Object byte[] $tailLength
    [void]$stream.Read($tailBytes, 0, $tailLength)
    $tailText = [System.Text.Encoding]::UTF8.GetString($tailBytes)
    $tailTrim = ($tailText.TrimEnd())
    return $tailTrim.EndsWith("]}")
  } finally {
    if ($stream) { $stream.Dispose() }
  }
}

$workingOutputGeoJson = $null
$tempDir = $null

try {
  Write-Step "Starting offline ZIP -> GeoJSON conversion"

  if (-not $ZipPath) {
    $ZipPath = Show-OpenFileDialog -Title "Select Texas streets ZIP" -Filter "ZIP files (*.zip)|*.zip"
  }
  if (-not $ZipPath) {
    Write-Step "Canceled."
    exit 1
  }

  if (-not (Test-Path -LiteralPath $ZipPath)) {
    throw "ZIP file not found: $ZipPath"
  }

  $defaultOut = Join-Path (Split-Path -Parent $ZipPath) (([System.IO.Path]::GetFileNameWithoutExtension($ZipPath)) + "-roads.geojson")
  if (-not $OutputGeoJson) {
    $OutputGeoJson = $defaultOut
    Write-Step "No output path selected. Using default: $OutputGeoJson"
  }
  $workingOutputGeoJson = "$OutputGeoJson.partial"
  if (Test-Path -LiteralPath $workingOutputGeoJson) {
    Remove-Item -LiteralPath $workingOutputGeoJson -Force -ErrorAction SilentlyContinue
  }

  $tempDir = New-WorkingTempDirectory -OutputPath $OutputGeoJson
  Write-Step "Temporary working folder: $tempDir"
  Write-Step "Extracting roads shapefile from ZIP..."
  $shpPath = Extract-RoadShapefileFromZip -ZipPathValue $ZipPath -DestinationRoot $tempDir

  $ogr = Get-Command ogr2ogr -ErrorAction SilentlyContinue
  if ($ogr) {
    Write-Step "Using GDAL ogr2ogr for fast conversion..."
    & $ogr.Source -f GeoJSON $workingOutputGeoJson $shpPath -skipfailures -lco RFC7946=YES
    if ($LASTEXITCODE -ne 0) {
      throw "ogr2ogr conversion failed with exit code $LASTEXITCODE."
    }
    if (-not (Test-Path -LiteralPath $workingOutputGeoJson)) {
      throw "Conversion finished but output file was not created: $workingOutputGeoJson"
    }
    if (-not (Test-GeoJsonFileLooksComplete -GeoJsonPath $workingOutputGeoJson)) {
      throw "Converted GeoJSON appears incomplete or corrupt: $workingOutputGeoJson"
    }
    Move-Item -LiteralPath $workingOutputGeoJson -Destination $OutputGeoJson -Force
    Write-Step "Conversion complete: $OutputGeoJson"
    Start-Process explorer.exe "/select,`"$OutputGeoJson`""
    Write-Step "You can now load this GeoJSON in TDS PAK with 'Load File'."
    if ($tempDir -and (Test-Path -LiteralPath $tempDir)) {
      Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    exit 0
  }

  $converted = $false
  $pythonInvocation = Get-PythonInvocation
  if ($pythonInvocation) {
    try {
      $pyScriptPath = Join-Path $tempDir "convert_roads.py"
      Set-Content -LiteralPath $pyScriptPath -Value (Build-PythonConverterScript) -Encoding UTF8
      Ensure-PyShpInstalled -PythonInvocation $pythonInvocation
      Write-Step "Converting shapefile to GeoJSON using Python..."
      $pyCmd = [string]$pythonInvocation.Command
      $pyPrefix = @($pythonInvocation.PrefixArgs)
      & $pyCmd @($pyPrefix + @($pyScriptPath, $shpPath, $workingOutputGeoJson))
      if ($LASTEXITCODE -ne 0) {
        throw "Python conversion failed with exit code $LASTEXITCODE."
      }
      if (-not (Test-Path -LiteralPath $workingOutputGeoJson)) {
        throw "Python conversion finished but output file was not created: $workingOutputGeoJson"
      }
      $converted = $true
    } catch {
      Write-Step ("Python path failed: " + $_.Exception.Message)
      Write-Step "Trying Node.js fallback..."
    }
  } else {
    Write-Step "Python runtime not detected. Trying Node.js fallback..."
  }

  if (-not $converted) {
    $nodeCmd = Get-NodeCommand
    if ($nodeCmd) {
      try {
        $nodeScriptPath = Join-Path $tempDir "convert_roads_node.js"
        Set-Content -LiteralPath $nodeScriptPath -Value (Build-NodeConverterScript) -Encoding UTF8
        Write-Step "Converting shapefile to GeoJSON using Node.js..."
        & $nodeCmd $nodeScriptPath $shpPath $workingOutputGeoJson
        if ($LASTEXITCODE -ne 0) {
          throw "Node.js conversion failed with exit code $LASTEXITCODE."
        }
        if (-not (Test-Path -LiteralPath $workingOutputGeoJson)) {
          throw "Node.js conversion finished but output file was not created: $workingOutputGeoJson"
        }
        $converted = $true
      } catch {
        Write-Step ("Node.js path failed: " + $_.Exception.Message)
        Write-Step "Trying built-in PowerShell fallback..."
      }
    } else {
      Write-Step "Node.js runtime not detected. Trying built-in PowerShell fallback..."
    }
  }

  if (-not $converted) {
    Write-Step "Converting shapefile to GeoJSON using built-in PowerShell parser..."
    Convert-ShapefileToGeoJsonPowerShell -ShpPath $shpPath -OutputGeoJson $workingOutputGeoJson
    if (-not (Test-Path -LiteralPath $workingOutputGeoJson)) {
      throw "PowerShell conversion finished but output file was not created: $workingOutputGeoJson"
    }
    $converted = $true
  }

  if (-not (Test-GeoJsonFileLooksComplete -GeoJsonPath $workingOutputGeoJson)) {
    throw "Converted GeoJSON appears incomplete or corrupt: $workingOutputGeoJson"
  }

  Move-Item -LiteralPath $workingOutputGeoJson -Destination $OutputGeoJson -Force

  Write-Step "Conversion complete: $OutputGeoJson"
  Start-Process explorer.exe "/select,`"$OutputGeoJson`""
  Write-Step "You can now load this GeoJSON in TDS PAK with 'Load File'."
  if ($tempDir -and (Test-Path -LiteralPath $tempDir)) {
    Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
  }
}
catch {
  if ($workingOutputGeoJson -and (Test-Path -LiteralPath $workingOutputGeoJson)) {
    Remove-Item -LiteralPath $workingOutputGeoJson -Force -ErrorAction SilentlyContinue
  }
  if ($tempDir -and (Test-Path -LiteralPath $tempDir)) {
    Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
  }
  Write-Host ""
  Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host ""
  Write-Host "Tips:"
  Write-Host "1) Ensure the drive used for TEMP/output has enough free space."
  Write-Host "2) Retry and select the ZIP from 'Download Texas Streets'."
  Write-Host "3) If Python is not installed, the built-in fallback can still work but may be slower."
  Write-Host "4) Optional speed boost: install Python (https://www.python.org/downloads/), GDAL, or Node.js (https://nodejs.org/)."
  Write-Host "5) If Windows shows Microsoft Store alias errors for python/py, disable those aliases or install Python from python.org."
  exit 1
}
