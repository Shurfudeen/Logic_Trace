import re
import csv
from pathlib import Path
from collections import defaultdict, deque


# ============================================================
# 1. CONFIGURATION
# ============================================================

# Known internal tags that should not be final initiators.
MANUAL_INTERNAL_TAGS = {
    "m2971HS101P_25",
}

# Constants that should not be reported as initiators.
CONSTANTS = {
    "TRUE", "FALSE", "0", "1", "0.0", "1.0",
    "T#5S", "T#0S", "T#1S",
}

MAX_TRACE_STEPS = 100000


# ============================================================
# 2. PARSE ONE STF FILE
# ============================================================

def parse_stf_file(file_path):
    """Extract variables, function blocks and ARC connections from one STF file."""

    file_path = Path(file_path)
    pou_name = file_path.stem

    variables = {}
    boxes = {}
    arcs = []
    parse_warnings = []

    with file_path.open("r", encoding="utf-8", errors="replace") as file:
        for line in file:
            line = line.strip()

            # Parse @VAR records.
            if line.startswith("@VAR:"):
                match = re.match(r"@VAR:(\d+).*?X=(.*)$", line)
                if match:
                    var_id = match.group(1)
                    tag = match.group(2).strip()
                    if var_id in variables:
                        parse_warnings.append(
                            f"Duplicate variable ID {var_id}; later record overrides earlier record"
                        )
                    if not tag:
                        parse_warnings.append(f"Variable {var_id} has an empty tag")
                    variables[var_id] = tag
                else:
                    parse_warnings.append(f"Could not parse @VAR record: {line}")

            # Parse @BOX records.
            elif line.startswith("@BOX:"):
                match = re.match(r"@BOX:(\d+).*?X=([^,\s]+)", line)
                if match:
                    box_id = match.group(1)
                    block_name = match.group(2).strip()
                    if box_id in boxes:
                        parse_warnings.append(
                            f"Duplicate block ID {box_id}; later record overrides earlier record"
                        )
                    boxes[box_id] = block_name
                else:
                    parse_warnings.append(f"Could not parse @BOX record: {line}")

            # Parse both normal @ARC connections and @NOT connections.
            # Both records define a directed connection in the FBD graph;
            # retain the record type so the trace can show where inversion occurs.
            elif line.startswith(("@ARC:", "@NOT:")):
                connection_type = "NOT" if line.startswith("@NOT:") else "ARC"
                match = re.search(
                    r"F=\((\d+),(\d+)\),T=\((\d+),(\d+)\)", line
                )
                if match:
                    from_id, from_port, to_id, to_port = match.groups()
                    arcs.append({
                        "from_id": from_id,
                        "from_port": from_port,
                        "to_id": to_id,
                        "to_port": to_port,
                        "connection_type": connection_type,
                    })
                else:
                    parse_warnings.append(
                        f"Could not parse @{connection_type} record: {line}"
                    )

    return {
        "pou": pou_name,
        "variables": variables,
        "boxes": boxes,
        "arcs": arcs,
        "parse_warnings": parse_warnings,
    }


# ============================================================
# 3. AUTOMATICALLY DISCOVER STF FILES AND GET TARGET TAG
# ============================================================

def discover_stf_files(root_folder, max_subfolder_levels=2):
    r"""Find STF files in the selected folder and up to N folder levels below it.

    For example, if the user enters:
        C:\RS-Projects\SAHIL\SCS1031

    this searches that folder, its immediate subfolders, and subfolders
    one level further down (two additional folder levels in total).
    """
    root_folder = Path(root_folder).resolve()
    discovered_files = []

    for file_path in root_folder.rglob("*"):
        if not file_path.is_file() or file_path.suffix.lower() != ".stf":
            continue

        # Count directory levels below the user-selected root. The file's
        # own filename is excluded from this count.
        relative_parent = file_path.parent.relative_to(root_folder)
        folder_depth = len(relative_parent.parts)

        if folder_depth <= max_subfolder_levels:
            discovered_files.append(file_path)

    return sorted(discovered_files)


def get_user_inputs():
    print("=" * 65)
    print("MULTI-POU FBD INPUT INITIATOR TRACER")
    print("=" * 65)
    print("Enter the project folder. ProTrace-RS will search the selected folder")
    print("and exactly up to 2 additional subfolder levels for .STF files.")

    while True:
        folder_input = input(
            "\nEnter the project folder path: "
        ).strip().strip('"').strip()

        folder_path = Path(folder_input)

        if not folder_path.exists():
            print("ERROR: Folder does not exist.")
            continue

        if not folder_path.is_dir():
            print("ERROR: The path is not a folder.")
            continue

        stf_files = discover_stf_files(folder_path)

        if not stf_files:
            print("ERROR: No .STF files were found within the selected folder and 2 additional folder levels.")
            continue

        break

    # Show the actual directories in which STF files were discovered.
    stf_directories = sorted({file_path.parent for file_path in stf_files})
    print(f"\nSTF files found: {len(stf_files)}")
    print(f"Folders containing STF files: {len(stf_directories)}")
    for directory in stf_directories:
        files_in_directory = [f for f in stf_files if f.parent == directory]
        print(f"\n  Folder: {directory}")
        for file_path in files_in_directory:
            print(f"    {file_path.name}")

    # Return only the project information. Target tags are entered later
    # in the repeated search loop, after each previous trace is completed.
    return folder_path, stf_files



# ============================================================
# 3A. LOAD AND CROSS-REFERENCE THE PROJECT TAG DICTIONARY
# ============================================================

def _normalise_column_name(value):
    """Normalise CSV headers so Tag Name, tag_name and TAGNAME are equivalent."""
    return re.sub(r"[^a-z0-9]", "", str(value).casefold())


def _read_csv_rows(csv_path):
    """Read Excel-exported CSV files with common encodings and delimiters."""
    last_error = None

    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with csv_path.open("r", encoding=encoding, newline="") as file:
                sample = file.read(16384)
                file.seek(0)

                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                except csv.Error:
                    # Excel CSV is normally comma-separated. Fall back to
                    # semicolon when it is clearly the more common separator.
                    delimiter = ";" if sample.count(";") > sample.count(",") else ","
                    dialect = csv.excel
                    dialect.delimiter = delimiter

                reader = csv.DictReader(file, dialect=dialect)
                if not reader.fieldnames:
                    raise ValueError("CSV has no header row.")

                headers = [str(h).strip().lstrip("\ufeff") for h in reader.fieldnames]
                rows = []
                for row in reader:
                    cleaned = {}
                    for original, value in row.items():
                        if original is not None:
                            cleaned[str(original).strip().lstrip("\ufeff")] = (
                                value.strip() if isinstance(value, str) else value
                            )
                    rows.append(cleaned)

                return headers, rows

        except (UnicodeError, csv.Error, OSError, ValueError) as error:
            last_error = error

    raise ValueError(f"Could not read CSV file {csv_path.name}: {last_error}")


def _identify_dictionary_columns(headers):
    """Identify the fixed project dictionary columns: Name, Data Type and Comment."""
    lookup = {_normalise_column_name(h): h for h in headers}

    tag_candidates = (
        "tag", "tagname", "tagid", "signal", "signalname",
        "variablename", "variable", "symbol", "instrumenttag",
        "name", "objectname", "fullname", "pointname"
    )
    tag_column = next(
        (lookup[name] for name in tag_candidates if name in lookup),
        None
    )
    comments_column = lookup.get("comment") or lookup.get("comments")
    datatype_column = lookup.get("datatype")

    # The supplied project dictionary format uses:
    # Name = tag, Data Type = datatype, Comment = tag description.
    # Keep a fallback for the older Comments spelling.
    # If the tag header is not recognised but the metadata columns are,
    # use the only remaining column as the tag column.
    if tag_column is None and comments_column and datatype_column:
        excluded = {comments_column, datatype_column}
        remaining = [h for h in headers if h not in excluded]
        if len(remaining) == 1:
            tag_column = remaining[0]

    return tag_column, comments_column, datatype_column


def _find_dictionary_csvs(project_folder, max_subfolder_levels=2):
    """Find CSV files in the selected project folder and up to 2 levels below."""
    project_folder = Path(project_folder).resolve()
    candidates = []

    for file_path in project_folder.rglob("*"):
        if not file_path.is_file() or file_path.suffix.casefold() != ".csv":
            continue

        depth = len(file_path.parent.relative_to(project_folder).parts)
        if depth > max_subfolder_levels:
            continue

        try:
            headers, _ = _read_csv_rows(file_path)
            tag_col, comments_col, datatype_col = _identify_dictionary_columns(headers)

            # A valid dictionary must expose tag, Comments and Datatype columns.
            if tag_col and comments_col and datatype_col:
                candidates.append(file_path)
        except (OSError, UnicodeError, csv.Error, ValueError):
            continue

    return sorted(candidates)


def load_project_dictionary(csv_path):
    """Load tag descriptions and data types from the selected dictionary CSV."""
    csv_path = Path(csv_path)
    if not csv_path.exists() or not csv_path.is_file():
        raise FileNotFoundError(f"Project dictionary file not found: {csv_path}")

    headers, rows = _read_csv_rows(csv_path)
    tag_column, comments_column, datatype_column = _identify_dictionary_columns(headers)

    missing = []
    if not tag_column:
        missing.append("tag-name")
    if not comments_column:
        missing.append("Comments")
    if not datatype_column:
        missing.append("Datatype")

    if missing:
        raise ValueError(
            f"{csv_path.name} is not in the expected project dictionary format. "
            "Expected columns: Name, Data Type, Comment. "
            + f"Headers found: {', '.join(headers)}"
        )

    dictionary = {}
    duplicate_count = 0

    for row in rows:
        raw_tag = (row.get(tag_column) or "").strip()
        if not raw_tag:
            continue

        key = raw_tag.casefold()
        if key in dictionary:
            duplicate_count += 1

        dictionary[key] = {
            "dictionary_tag": raw_tag,
            "description": (row.get(comments_column) or "").strip(),
            "datatype": (row.get(datatype_column) or "").strip(),
        }

    if not dictionary:
        raise ValueError(f"No tag records were found in {csv_path.name}.")

    print(f"Project dictionary selected: {csv_path}")
    print(f"Project dictionary loaded: {len(dictionary)} unique tags")
    print(f"Tag column: {tag_column} | Comments column: {comments_column} | Datatype column: {datatype_column}")

    if duplicate_count:
        print(
            f"WARNING: {duplicate_count} duplicate tag row(s) found; "
            "the last occurrence is used."
        )

    return dictionary


def get_dictionary_input(project_folder):
    """Automatically locate and load the project dictionary CSV from the project folder."""
    print("\nSearching for the project dictionary CSV in the selected project folder...")
    candidates = _find_dictionary_csvs(project_folder)

    if not candidates:
        raise FileNotFoundError(
            "No suitable project dictionary CSV was found in the selected folder "
            "or its first 2 subfolder levels. The CSV must contain a tag-name "
            "column, a Comment column and a Data Type column."
        )

    if len(candidates) == 1:
        selected_csv = candidates[0]
    else:
        # Prefer an unambiguous filename containing 'dictionary'.
        named_matches = [
            path for path in candidates
            if "dictionary" in path.stem.casefold()
            or "taglist" in path.stem.casefold()
            or "tag_list" in path.stem.casefold()
        ]

        if len(named_matches) == 1:
            selected_csv = named_matches[0]
        else:
            print("\nMore than one suitable dictionary CSV was found:")
            for index, path in enumerate(candidates, start=1):
                print(f"{index}. {path}")
            while True:
                choice = input("Select the project dictionary number: ").strip()
                if choice.isdigit() and 1 <= int(choice) <= len(candidates):
                    selected_csv = candidates[int(choice) - 1]
                    break
                print("Invalid selection. Enter one of the listed numbers.")

    return load_project_dictionary(selected_csv)


def get_dictionary_record(tag, project_dictionary):
    """Find dictionary metadata using case-insensitive matching.

    Exact tag matches take priority. If no exact match exists, a single
    optional .v / no-.v counterpart is accepted.
    """
    if not project_dictionary:
        return None

    key = tag.strip().casefold()
    if key in project_dictionary:
        return project_dictionary[key]

    alternate = key[:-2] if key.endswith(".v") else key + ".v"
    if alternate in project_dictionary:
        return project_dictionary[alternate]

    return None




def is_reset_ready_internal_tag(tag, project_dictionary, internal_tags):
    """Identify internal tags marked READY TO RESET, RDY or RST in the dictionary."""
    internal_lookup = {item.casefold() for item in internal_tags}
    if tag.strip().casefold() not in internal_lookup:
        return False

    record = get_dictionary_record(tag, project_dictionary)
    if not record:
        return False

    description = re.sub(r"\s+", " ", record.get("description", "")).strip()
    return bool(re.search(
        r"\b(?:ready\s+to\s+reset|rdy|rst)\b",
        description,
        re.IGNORECASE
    ))



# ============================================================
# 4. LOAD ALL STF FILES
# ============================================================

def load_all_pous(stf_files):
    """Load the exact STF file list discovered from the selected project folder."""
    stf_files = sorted(Path(file_path) for file_path in stf_files)

    if not stf_files:
        raise FileNotFoundError("No STF files were supplied for loading.")

    pous = {}
    duplicate_pou_names = []
    for file_path in stf_files:
        pou = parse_stf_file(file_path)
        if pou["pou"] in pous:
            duplicate_pou_names.append(
                f"{file_path} shares POU name {pou['pou']!r} with another STF file"
            )
        pous[pou["pou"]] = pou

    if duplicate_pou_names:
        raise ValueError(
            "Duplicate POU names detected; node IDs would collide:\n  "
            + "\n  ".join(duplicate_pou_names)
        )

    return pous


# ============================================================
# 5. BUILD THE CONNECTION DATABASE
# ============================================================

def build_database(pous):
    tags_to_nodes = defaultdict(list)
    node_to_tag = {}
    node_to_block = {}
    incoming = defaultdict(list)
    outgoing = defaultdict(list)
    dangling_arcs = []

    # Register variables and blocks. Also record which POUs contain each tag.
    tag_pous = defaultdict(set)
    for pou_name, pou in pous.items():
        for var_id, tag in pou["variables"].items():
            node = (pou_name, var_id, "VAR")
            tags_to_nodes[tag].append(node)
            node_to_tag[node] = tag
            tag_pous[tag].add(pou_name)

        for box_id, block_name in pou["boxes"].items():
            node = (pou_name, box_id, "BOX")
            node_to_block[node] = block_name

    # Build directed connections within each POU.
    # Treat @NOT exactly like @ARC for graph traversal. The connection type
    # is retained as metadata so the trace log can identify inverted signals.
    accepted_not_count = 0
    for pou_name, pou in pous.items():
        for arc in pou["arcs"]:
            from_id = arc["from_id"]
            to_id = arc["to_id"]
            connection_type = arc.get("connection_type", "ARC")

            if connection_type == "NOT":
                accepted_not_count += 1

            if from_id in pou["variables"]:
                source = (pou_name, from_id, "VAR")
            elif from_id in pou["boxes"]:
                source = (pou_name, from_id, "BOX")
            else:
                dangling_arcs.append(
                    f"{pou_name}: {connection_type} source ID {from_id} is neither a VAR nor BOX"
                )
                continue

            if to_id in pou["variables"]:
                destination = (pou_name, to_id, "VAR")
            elif to_id in pou["boxes"]:
                destination = (pou_name, to_id, "BOX")
            else:
                dangling_arcs.append(
                    f"{pou_name}: {connection_type} destination ID {to_id} is neither a VAR nor BOX"
                )
                continue

            connection = {
                "source": source,
                "destination": destination,
                "from_port": arc["from_port"],
                "to_port": arc["to_port"],
                "connection_type": connection_type,
            }
            outgoing[source].append(connection)
            incoming[destination].append(connection)

    # Detect intermediate/branch-through tags using @ARC records only.
    # @NOT must not influence this heuristic, or it can incorrectly classify
    # unrelated field inputs as internal tags (or the reverse).
    tags_from_side = set()
    tags_to_side = set()
    for pou in pous.values():
        for arc in pou["arcs"]:
            if arc.get("connection_type", "ARC") != "ARC":
                continue
            from_id = arc["from_id"]
            to_id = arc["to_id"]
            if from_id in pou["variables"]:
                tags_from_side.add(pou["variables"][from_id])
            if to_id in pou["variables"]:
                tags_to_side.add(pou["variables"][to_id])

    automatic_internal_tags = tags_from_side & tags_to_side

    # Index block instances by their full instance names, across all POUs.
    # Example: _1471XFD527W23FB -> its BOX node in POU0001.
    block_names_to_nodes = defaultdict(list)
    for node, block_name in node_to_block.items():
        block_names_to_nodes[block_name].append(node)

    return {
        "tags_to_nodes": tags_to_nodes,
        "node_to_tag": node_to_tag,
        "node_to_block": node_to_block,
        "incoming": incoming,
        "outgoing": outgoing,
        "automatic_internal_tags": automatic_internal_tags,
        "block_names_to_nodes": block_names_to_nodes,
        "dangling_arcs": dangling_arcs,
    }


# ============================================================
# 6. IDENTIFY CONSTANTS
# ============================================================

def is_quoted_description(tag):
    """Identify descriptive text enclosed completely in single quotes."""
    cleaned = tag.strip()
    return len(cleaned) >= 2 and cleaned.startswith("'") and cleaned.endswith("'")


def is_constant(tag):
    cleaned = tag.strip()

    if cleaned.upper() in CONSTANTS:
        return True

    if re.fullmatch(r"-?\d+(?:\.\d+)?", cleaned):
        return True

    if re.fullmatch(
        r"T#(?:\d+H)?(?:\d+M)?(?:\d+(?:\.\d+)?S)?(?:\d+MS)?",
        cleaned,
        flags=re.IGNORECASE,
    ):
        return True

    return False


# ============================================================
# 7. IDENTIFY HARDWIRED INPUT INDICATORS
# ============================================================

def is_hardwired_input(tag):
    lowered = tag.lower()
    return lowered.endswith(".v") or lowered.endswith(".status")


# ============================================================
# 8. TRACE BACKWARD FROM THE TARGET
# ============================================================

def resolve_parent_block(tag, block_names_to_nodes):
    """
    Resolve a dotted signal tag to its function-block instance.

    Prefer the longest exact instance-name prefix. Some STF signal names
    include multi-part attributes such as '.ANLGS.NLTR'; splitting only at
    the final dot incorrectly produces a parent such as '...ANLGS'.
    """
    # First handle the known compound attribute suffix explicitly.
    compound_suffixes = (".ANLGS.NLTR",)
    for suffix in compound_suffixes:
        if tag.endswith(suffix):
            base_name = tag[:-len(suffix)]
            if base_name in block_names_to_nodes:
                return base_name, block_names_to_nodes[base_name], suffix.lstrip(".")

    # General case: any registered block name that is a full prefix of the tag.
    candidates = [
        name for name in block_names_to_nodes
        if tag.startswith(name + ".")
    ]
    if candidates:
        parent_name = max(candidates, key=len)
        output_name = tag[len(parent_name) + 1:]
        return parent_name, block_names_to_nodes[parent_name], output_name

    # No registered parent block was found.
    if "." in tag:
        parent_name, output_name = tag.rsplit(".", 1)
        return parent_name, [], output_name
    return None, [], None

def trace_target(target_tag, database, project_dictionary=None):
    tags_to_nodes = database["tags_to_nodes"]
    node_to_tag = database["node_to_tag"]
    node_to_block = database["node_to_block"]
    incoming = database["incoming"]
    block_names_to_nodes = database["block_names_to_nodes"]

    internal_tags = (
        database["automatic_internal_tags"] | MANUAL_INTERNAL_TAGS
    )

    # The user normally enters the base tag without ".v". Check both
    # the entered spelling and its possible ".v" form before tracing.
    # Exact matches take priority; case-insensitive matching is accepted
    # only when it identifies one unique candidate.
    target_candidates = [target_tag]
    if not target_tag.lower().endswith(".v"):
        target_candidates.append(target_tag + ".v")

    resolved_tag = next(
        (candidate for candidate in target_candidates
         if candidate in tags_to_nodes),
        None
    )

    if resolved_tag is None:
        for candidate in target_candidates:
            similar_tags = sorted(
                tag for tag in tags_to_nodes
                if tag.casefold() == candidate.casefold()
            )
            if len(similar_tags) == 1:
                resolved_tag = similar_tags[0]
                break
            if len(similar_tags) > 1:
                print(f"\nERROR: Ambiguous case-insensitive target: {candidate}")
                for tag in similar_tags:
                    print(f"  {tag}")
                return [], [], []

    if resolved_tag is None:
        checked_tag = target_tag if target_tag.lower().endswith(".v") else target_tag + ".v"
        print(f"\nERROR: Target tag not found: {target_tag} (also checked: {checked_tag})")
        return [], [], []

    if resolved_tag != target_tag:
        print(f"\nNOTICE: Target resolved: {target_tag!r} -> {resolved_tag!r}")
    target_tag = resolved_tag
    start_nodes = tags_to_nodes.get(resolved_tag, [])

    queue = deque(start_nodes)
    visited = set()
    initiators = set()
    trace_log = []
    traced_pous = set()
    steps = 0

    while queue:
        node = queue.popleft()
        if node in visited:
            continue

        visited.add(node)
        steps += 1
        if steps > MAX_TRACE_STEPS:
            trace_log.append("STOPPED: maximum trace steps reached.")
            break

        pou_name, item_id, item_type = node
        traced_pous.add(pou_name)

        # CASE A: A function-block instance. Trace its connected inputs.
        if item_type == "BOX":
            block_name = node_to_block.get(node, "UNKNOWN")
            trace_log.append(
                f"BLOCK: {pou_name} / {block_name} / ID {item_id}"
            )

            block_inputs = incoming.get(node, [])
            if not block_inputs:
                trace_log.append(
                    f"  SOURCE-LESS TERMINAL BLOCK: {block_name} has no incoming "
                    f"ARC/@NOT connections; reporting it as a final initiator"
                )
                # A source-less block reached during backward tracing is a valid
                # terminal result in its own right. Keep it in the same final
                # initiator list as source variables, but do not report placeholders.
                if block_name and block_name != "UNKNOWN":
                    initiators.add(block_name)
                else:
                    trace_log.append(
                        "  SOURCE-LESS BLOCK NOT REPORTED: block name is unresolved"
                    )

            for connection in block_inputs:
                source = connection["source"]
                connection_type = connection.get("connection_type", "ARC")
                polarity_note = " (inverted @NOT connection)" if connection_type == "NOT" else ""
                trace_log.append(
                    f"  Connected input [{connection_type}]: "
                    f"{source[0]} / {source[2]} / ID {source[1]}{polarity_note}"
                )
                queue.append(source)
            continue

        # CASE B: A variable.
        tag = node_to_tag[node]

        # Ignore this branch when an internal tag is marked READY TO RESET,
        # RDY or RST in the project dictionary. Other queued branches continue.
        if is_reset_ready_internal_tag(tag, project_dictionary, internal_tags):
            record = get_dictionary_record(tag, project_dictionary)
            description = record.get("description", "") if record else ""
            trace_log.append(
                f"RESET-READY INTERNAL PATH IGNORED: {tag} "
                f"[description: {description}] [{pou_name}]"
            )
            continue


        if not tag.strip() or tag.strip().strip("\"'").strip() == "":
            trace_log.append(f"EMPTY TAG IGNORED [{pou_name}, VAR ID {item_id}]")
            continue

        if is_constant(tag):
            trace_log.append(f"CONSTANT IGNORED: {tag} [{pou_name}]")
            continue

        # Ignore descriptions enclosed in single quotes. Keep their VAR
        # records in the database so graph connection IDs remain valid.
        if is_quoted_description(tag):
            trace_log.append(f"DESCRIPTION IGNORED: {tag} [{pou_name}]")
            continue

        # Check local connections before treating .v/.status as a final source.
        # In STF files, a tag ending in .v or .status can also appear on the
        # OUTPUT side of a function block. If it has an incoming ARC, trace
        # upstream first rather than reporting it as a final initiator.
        variable_incoming = incoming.get(node, [])

        if is_hardwired_input(tag):
            if variable_incoming:
                trace_log.append(
                    f"HARDWIRED-SUFFIX TAG IS AN OUTPUT/CONNECTED VARIABLE: "
                    f"{tag}; tracing upstream in {pou_name}"
                )
                for connection in variable_incoming:
                    queue.append(connection["source"])
                continue

            # A tag such as INSTANCE.v may be a block output even when its
            # output variable has no local incoming ARC in this POU. Try the
            # matching parent block name before treating it as a source.
            if tag.lower().endswith(".v"):
                parent_block_name = tag[:-2]
                parent_block_nodes = block_names_to_nodes.get(
                    parent_block_name, []
                )
                if parent_block_nodes:
                    trace_log.append(
                        f"BLOCK OUTPUT: {tag} -> parent block "
                        f"{parent_block_name} (output: v)"
                    )
                    for parent_node in parent_block_nodes:
                        queue.append(parent_node)
                    continue

            # A final .v/.status source must be on the input side of a block:
            # the variable must feed a BOX through an outgoing ARC.
            feeds_block_input = any(
                connection["destination"][2] == "BOX"
                for connection in database["outgoing"].get(node, [])
            )

            if feeds_block_input:
                if tag == target_tag:
                    trace_log.append(
                        f"TARGET TAG NOT REPORTED AS ITS OWN INITIATOR: {tag} [{pou_name}]"
                    )
                else:
                    initiators.add(tag)
                    trace_log.append(
                        f"INITIATOR (HARDWIRED INPUT): {tag} [{pou_name}]"
                    )
            else:
                trace_log.append(
                    f"NOT A FINAL INITIATOR: {tag} [{pou_name}] "
                    f"is not confirmed on a block input side"
                )
            continue

        # A dotted tag other than .v/.status may represent a block output.
        # Resolve the portion before the final dot as the parent block instance.
        # Example: _1471XFD527W23FB.HTRP -> _1471XFD527W23FB.
        if "." in tag:
            parent_block_name, parent_block_nodes, output_name = resolve_parent_block(
                tag, block_names_to_nodes
            )

            if parent_block_nodes:
                trace_log.append(
                    f"BLOCK OUTPUT: {tag} -> parent block "
                    f"{parent_block_name} (output: {output_name})"
                )
                for parent_node in parent_block_nodes:
                    queue.append(parent_node)
                continue

            # For compound analog-status suffixes, do not misreport the
            # suffixed signal as a field initiator when its parent block
            # is absent from the parsed @BOX records. Keep the trace explicit
            # so a missing/renamed block can be investigated.
            if tag.endswith(".ANLGS.NLTR"):
                trace_log.append(
                    f"UNRESOLVED COMPOUND BLOCK OUTPUT: {tag}; expected parent "
                    f"{tag[:-len('.ANLGS.NLTR')]} was not found in @BOX records; "
                    "not reported as an initiator"
                )
                continue

            # No matching block instance was found. Continue with normal
            # incoming/cross-POU checks instead of prematurely assuming a block.
            trace_log.append(
                f"BLOCK OUTPUT LOOKUP: no block named {parent_block_name} "
                f"found for {tag}; checking other connections"
            )

        # If this variable has a local incoming connection, follow it upstream.
        if variable_incoming:
            trace_log.append(
                f"VARIABLE: {tag} has an incoming connection in {pou_name}"
            )
            for connection in variable_incoming:
                if connection.get("connection_type") == "NOT":
                    trace_log.append(
                        f"  Following inverted @NOT connection from "
                        f"{connection['source'][0]} / {connection['source'][2]} / "
                        f"ID {connection['source'][1]}"
                    )
                queue.append(connection["source"])
            continue

        # Link the same tag across different POUs when another occurrence
        # has incoming connections.
        matching_nodes = tags_to_nodes.get(tag, [])
        alternative_nodes = [
            other_node for other_node in matching_nodes
            if other_node != node and incoming.get(other_node)
        ]

        if alternative_nodes:
            trace_log.append(
                f"CROSS-POU LINK: {tag} -> {len(alternative_nodes)} candidate node(s)"
            )
            if len(alternative_nodes) > 1:
                trace_log.append(
                    f"  AMBIGUITY: multiple upstream definitions exist for {tag}; "
                    "all candidates are traced, so verify these links in the source project"
                )
            for other_node in alternative_nodes:
                queue.append(other_node)
            continue

        # Leaf variables marked as internal should not be final initiators.
        if tag in internal_tags:
            trace_log.append(f"EXCLUDED INTERNAL TAG: {tag}")
            continue

        # Only an unresolved leaf with no recognized parent is reported here.
        if tag == target_tag:
            trace_log.append(f"TARGET TAG NOT REPORTED AS ITS OWN INITIATOR: {tag} [{pou_name}]")
            continue
        initiators.add(tag)
        trace_log.append(f"INITIATOR: {tag} [{pou_name}]")

    return sorted(initiators), sorted(traced_pous), trace_log


# ============================================================
# 9. DISPLAY THE RESULTS
# ============================================================

def display_results(target_tag, initiators, traced_pous, project_dictionary):
    print("\n" + "=" * 100)
    print(f"RESULTS FOR: {target_tag}")
    print("=" * 100)

    print("\n1. POUs FOUND")
    if traced_pous:
        for number, pou_name in enumerate(traced_pous, start=1):
            print(f"{number}. {pou_name}")
    else:
        print("No POUs found.")

    print("\n2. INITIATORS IDENTIFIED")
    if initiators:
        print(
            f"{'#':<5} {'RESULT TAG':<38} {'DATATYPE':<18} "
            f"{'DICTIONARY':<14} DESCRIPTION"
        )
        print("-" * 120)

        for number, tag in enumerate(initiators, start=1):
            record = get_dictionary_record(tag, project_dictionary)

            if record:
                datatype = record["datatype"] or "Not specified"
                description = record["description"] or "Not specified"
                dictionary_status = "MATCHED"
            else:
                datatype = "Not found"
                description = "Not found in project dictionary"
                dictionary_status = "NOT FOUND"

            print(
                f"{number:<5} {tag:<38.38} {datatype:<18.18} "
                f"{dictionary_status:<14} {description}"
            )
    else:
        print("No initiators identified.")

    print("\nDictionary cross-verification:")
    print("MATCHED = tag exists in the project dictionary.")
    print("NOT FOUND = tag was identified by STF tracing but is absent from the dictionary.")
    print("=" * 100)


# ============================================================
# 10. MAIN PROGRAM
# ============================================================

def main():
    folder_path, stf_files = get_user_inputs()

    try:
        project_dictionary = get_dictionary_input(folder_path)
        print("\nLoading STF files...")
        pous = load_all_pous(stf_files)
        database = build_database(pous)

        # ------------------------------------------------------------
        # REPEATED TAG-TRACING SESSION
        # ------------------------------------------------------------
        search_count = 0

        while True:
            target_tag = input(
                "\nEnter the target tag to trace (e.g. _09SOV2003): "
            ).strip()

            while not target_tag:
                print("ERROR: Target tag cannot be empty.")
                target_tag = input(
                    "Enter the target tag to trace (e.g. _09SOV2003): "
                ).strip()

            search_count += 1
            print("\n" + "=" * 65)
            print(f"SEARCH NUMBER: {search_count}")
            print("=" * 65)

            # Each trace_target() call creates its own queue, visited set,
            # initiator set and trace log, so searches remain independent.
            initiators, traced_pous, _trace_log = trace_target(target_tag, database, project_dictionary)
            display_results(target_tag, initiators, traced_pous, project_dictionary)

            while True:
                choice = input(
                    "\nWould you like to search for another target tag? (Y/N): "
                ).strip().casefold()

                if choice in ("y", "yes"):
                    # Return to the target-tag prompt for the next search.
                    break
                elif choice in ("n", "no"):
                    print("\n" + "=" * 65)
                    print("SEARCH SESSION COMPLETED")
                    print(f"Total tag searches performed: {search_count}")
                    print("=" * 65)
                    return
                else:
                    print("Invalid entry. Please enter Y or N.")

    except (OSError, UnicodeError, ValueError, csv.Error) as error:
        print(f"\nERROR: Unable to process STF files: {error}")


# ============================================================
# 11. START PROGRAM
# ============================================================

if __name__ == "__main__":
    main()
