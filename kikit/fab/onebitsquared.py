import click
from pcbnewTransition import pcbnew
import csv
import configparser
import os
import re
import sys
import shutil
from pathlib import Path
from kikit.fab.common import *
from kikit.common import *
from kikit.export import gerberImpl, exportSettingsPcbway

def collectSolderTypes(board):
    result = {}
    for footprint in board.GetFootprints():
        if excludeFromPos(footprint):
            continue
        if hasNonSMDPins(footprint):
            result[footprint.GetReference()] = "thru-hole"
        else:
            result[footprint.GetReference()] = "SMD"

    return result

def addVirtualToRefsToIgnore(refsToIgnore, board):
    for footprint in board.GetFootprints():
        if excludeFromPos(footprint):
            refsToIgnore.append(footprint.GetReference())

def collectBom(components, keyNames, manufacturerFields, partNumberFields,
               descriptionFields, notesFields, typeFields, footprintFields,
               ignore):
    bom = {}

    # Use KiCad footprint as fallback for footprint
    footprintFields.append("Footprint")
    # Use value as fallback for description
    descriptionFields.append("Value")

    for c in components:
        if getUnit(c) != 1:
            continue
        reference = getReference(c)
        if reference.startswith("#PWR") or reference.startswith("#FL") or reference in ignore:
            continue
        if hasattr(c, "in_bom") and not c.in_bom:
            continue
        if hasattr(c, "on_board") and not c.on_board:
            continue
        if hasattr(c, "dnp") and c.dnp:
            continue

        key = None
        for keyName in keyNames:
            key = getField(c, keyName)
            if key is not None:
                break
        manufacturer = None
        for manufacturerName in manufacturerFields:
            manufacturer = getField(c, manufacturerName)
            if manufacturer is not None:
                break
        partNumber = None
        for partNumberName in partNumberFields:
            partNumber = getField(c, partNumberName)
            if partNumber is not None:
                break
        description = None
        for descriptionName in descriptionFields:
            description = getField(c, descriptionName)
            if description is not None:
                break
        notes = None
        for notesName in notesFields:
            notes = getField(c, notesName)
            if notes is not None:
                break
        solderType = None
        for typeName in typeFields:
            solderType = getField(c, typeName)
            if solderType is not None:
                break
        footprint = None
        for footprintName in footprintFields:
            footprint = getField(c, footprintName)
            if footprint is not None:
                break

        cType = (
            key,
            description,
            footprint,
            manufacturer,
            partNumber,
            notes,
            solderType
        )
        bom[cType] = bom.get(cType, []) + [reference]
    return bom

def bomToXsv(bomData, filename, nBoards, types, delim=','):
    with open(filename, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile, delimiter=delim)
        writer.writerow(["Item #", "Key", "Designator", "Qty", "Manufacturer",
                         "Mfg Part #", "Description / Value", "Footprint",
                         "Type", "Your Instructions / Notes"])
        item_no = 1

        tmp = {}
        for cType, references in bomData.items():
            tmp[references[0]] = (references, cType)

        for i in sorted(tmp, key=naturalComponentKey):
            references, cType = tmp[i]
            references = sorted(references, key=naturalComponentKey)
            key, description, footprint, manufacturer, partNumber, notes, solderType = cType
            if solderType is None:
                solderType = types[references[0]]
            writer.writerow([item_no, key, ",".join(references),
                             len(references) * nBoards, manufacturer,
                             partNumber, description, footprint,
                             solderType, notes])
            item_no += 1

def to_mm(nm):
    return nm/1e+6

def to_mm_coords(coords):
    mm_coords = ()
    for c in coords:
        mm_coords + (to_mm(c))
    return mm_coords

def generateHanwhaSSA(bomData, posData, panel_preset, board, filename, panel_side='T'):
    ref_v_key = dict()
    for cType, references in bomData.items():
        for ref in references:
            # print(f"ref {ref} type {cType[0]}")
            ref_v_key[ref] = cType[0]

    import kikit.panelize_ui_impl as ki
    from kikit.panelize import Panel

    if panel_side == 'T':
        panel_side_name = "TOP"
    elif panel_side == 'B':
        panel_side_name = "BOTTOM"
    else:
        print(f"The provided panel side \"{panel_side}\" is not valid. Only \"T\" and \"B\" are supported.")
        exit

    loadedBoard = pcbnew.LoadBoard(board)

    imageBoundingBox = findBoardBoundingBox(loadedBoard)
    imageOrigin = imageBoundingBox.GetOrigin()
    imageSize = imageBoundingBox.GetSize()

    print(f"image origin {imageOrigin}")
    print(f"image size {imageSize}")

    imageThickness = loadedBoard.GetDesignSettings().GetBoardThickness()
    print(f"image thickness {imageThickness}")

    if panel_preset["layout"]["type"] != "grid":
        print("ERROR: We only support grid panel layout")
        exit

    panel_array = (panel_preset["layout"]["cols"], panel_preset["layout"]["rows"])
    panel_gap = (panel_preset["layout"]["hspace"], panel_preset["layout"]["vspace"])

    if panel_preset["framing"]["type"] != "frame":
        print("ERROR: We only support panels with frames")
        exit

    panel_frame_gap = panel_preset["framing"]["space"]
    panel_frame_width = panel_preset["framing"]["width"]

    panel_size = ((panel_frame_width + panel_frame_gap) * 2 + \
                  imageSize[0] * panel_array[0] + panel_gap[0] * (panel_array[0] - 1), \
                  (panel_frame_width + panel_frame_gap) * 2 + \
                  imageSize[1] * panel_array[1] + panel_gap[1] * (panel_array[1] - 1))
    print(f"panel size {panel_size}")

    if panel_side == 'T':
        placement_origin = (-(panel_frame_width + panel_frame_gap), panel_frame_width + panel_frame_gap)
    else:
        placement_origin = (-(panel_frame_width + panel_frame_gap + imageSize[0]), panel_frame_width + panel_frame_gap)
    print(f"placement origin {placement_origin}")

    panel_array_offset = (imageSize[0] + panel_gap[0], imageSize[1] + panel_gap[1])
    print(f"image offset {panel_array_offset}")

    if panel_preset["fiducials"]["type"] != "3fid":
        print("ERROR: We only support panels with 3fid fiducials")
        exit

    fid_offset = (panel_preset["fiducials"]["hoffset"], panel_preset["fiducials"]["voffset"])
    if panel_side == 'T':
        #fiducials = ((panel_size[0] - fid_offset[0] - placement_origin[0], fid_offset[1] - placement_origin[1]), \
        #             (panel_size[0] - fid_offset[0] - placement_origin[0], panel_size[1] - fid_offset[1] - placement_origin[1]), \
        #             (fid_offset[0] - placement_origin[0], panel_size[1] - fid_offset[1] - placement_origin[1]))
        fiducials = ((placement_origin[0] + panel_size[0] - fid_offset[0], -placement_origin[1] + fid_offset[1]), \
                     (placement_origin[0] + panel_size[0] - fid_offset[0], -placement_origin[1] + panel_size[1] - fid_offset[1]), \
                     (placement_origin[0]                 + fid_offset[0], -placement_origin[1] + panel_size[1] - fid_offset[1]))
    else:
        fiducials = ((placement_origin[0]                 + fid_offset[0], -placement_origin[1] + fid_offset[1]), \
                     (placement_origin[0] + panel_size[0] - fid_offset[0], -placement_origin[1] + panel_size[1] - fid_offset[1]), \
                     (placement_origin[0]                 + fid_offset[0], -placement_origin[1] + panel_size[1] - fid_offset[1]))

    for fid in fiducials:
        print(f"fid {fid}")

    config = configparser.ConfigParser()
    config.add_section('VERSION')

    config.add_section('PCB')
    config['PCB']['Unit System'] = 'MILIMETER'
    config['PCB']['Coordinate'] = 'LOWER RIGHT'
    config['PCB']['Rotation'] = '0'
    config['PCB']['Placement Origin'] = f", {to_mm(placement_origin[0]):.3f}, {to_mm(placement_origin[1]):.3f}"
    config['PCB']['Fiducial'] = f"CIRCLE, {to_mm(fiducials[0][0]):.3f}, {to_mm(fiducials[0][1]):.3f}, {to_mm(fiducials[2][0]):.3f}, {to_mm(fiducials[2][1]):.3f}"
    config['PCB']['Accept Mark'] = 'NONE, 0, 0'
    config['PCB']['Bad Mark'] = 'NONE, 0, 0'

    config.add_section('BOARD')
    config['BOARD']['Board Name'] = panel_preset["text"]["text"] + " " + panel_side_name
    config['BOARD']['PCB Size'] = f"{to_mm(panel_size[0]):.3f}, {to_mm(panel_size[1]):.3f}, {to_mm(imageThickness):.3f}"
    config['BOARD']['Array'] = f"{panel_array[0]}, {panel_array[1]}, LOWER RIGHT"
    config['BOARD']['Array Offset'] = f"{to_mm(panel_array_offset[0]):.3f}, {to_mm(panel_array_offset[1]):.3f}"


    with open(filename, "w", newline="\r\n", encoding="utf-8") as configfile:
        config.write(configfile)
        configfile.write("[FIDUCIAL]\n")
        fid_id = 0
        for fid in fiducials:
            fid_id += 1
            configfile.write(f"{fid_id} CIRCLE,{to_mm(fid[0]):.3f},{to_mm(fid[1]):.3f}\n")
        configfile.write("\n[PLACEMENTS]\n")
        writer = csv.writer(configfile, delimiter=' ', quoting=csv.QUOTE_ALL, lineterminator='\n')
        # Placement columns
        # Ref. X Y Z T LF-Shape LF1-X LF1-Y LF2-X LF2-Y CM_No Skip P/N P/F D/C
        for line in sorted(posData, key=lambda x: naturalComponentKey(x[0])):
            ref, x, y, side, t = line
            if side != panel_side:
                continue
            if ref not in ref_v_key.keys():
                print(f"Skipping pos line for ref {ref} due to no key (panel side {panel_side})")
                continue
            if panel_side == 'T':
                line = list((ref, f"{-x:.3f}", f"{y:.3f}", '0.000', f"{t:.3f}", 'NONE', '0', '0', '0', '0', '1991', '0', ref_v_key[ref], "", ""))
            else:
                line = list((ref, f"{x:.3f}", f"{y:.3f}", '0.000', f"{t:.3f}", 'NONE', '0', '0', '0', '0', '1991', '0', ref_v_key[ref], "", ""))
            writer.writerow(line)


def exportOneBitSquared(preset, board, outputdir, schematic, nametemplate, drc):
    """
    Prepare fabrication files for 1BitSquared PCBA
    """

    # Old params
    ignore = ""
    key = "Key, 1b2-bom-key"
    manufacturer = ""
    partnumber = ""
    description = ""
    notes = ""
    soldertype = ""
    footprint = ""
    corrections = ""
    correctionpatterns = None
    missingkeyerror = True
    missingerror = False
    nboards = 1

    # print("Board: %s" % board)

    ensureValidBoard(board)
    loadedBoard = pcbnew.LoadBoard(board)

    if drc:
        ensurePassingDrc(loadedBoard)

    refsToIgnore = parseReferences(ignore)
    removeComponents(loadedBoard, refsToIgnore)
    Path(outputdir).mkdir(parents=True, exist_ok=True)

    gerberdir = os.path.join(outputdir, "gerber")
    shutil.rmtree(gerberdir, ignore_errors=True)
    gerberImpl(board, gerberdir, settings=exportSettingsPcbway)

    archiveName = expandNameTemplate(nametemplate, "gerbers", loadedBoard)
    shutil.make_archive(os.path.join(outputdir, archiveName), "zip", outputdir, "gerber")

    ensureValidSch(schematic)

    components = extractComponents(schematic)
    correctionFields    = [x.strip() for x in corrections.split(",")]
    keyFields           = [x.strip() for x in key.split(",")]
    manufacturerFields  = [x.strip() for x in manufacturer.split(",")]
    partNumberFields    = [x.strip() for x in partnumber.split(",")]
    descriptionFields   = [x.strip() for x in description.split(",")]
    notesFields         = [x.strip() for x in notes.split(",")]
    typeFields          = [x.strip() for x in soldertype.split(",")]
    footprintFields     = [x.strip() for x in footprint.split(",")]
    addVirtualToRefsToIgnore(refsToIgnore, loadedBoard)
    print("collecting bom")
    bom = collectBom(components, keyFields, manufacturerFields, partNumberFields,
                     descriptionFields, notesFields, typeFields,
                     footprintFields, refsToIgnore)

    print("missing fields")
    missingFields = False
    missingKeys = False
    for type, references in bom.items():
        key, _, _, manu, partno, _, _ = type
        if not key:
            missingKeys = True
            for r in references:
                print(f"WARNING: Component {r} is missing key")
        if not key or not manu or not partno:
            missingFields = True
            for r in references:
                print(f"WARNING: Component {r} is missing key, manufacturer and/or part number")
    if missingFields and missingerror:
        sys.exit("There are components with missing ordercode, aborting")
    if missingKeys and missingkeyerror:
        sys.exit("There are components with missing ordercode, aborting")

    print("Collecting pos data")
    posData = collectPosData(loadedBoard, correctionFields, posFilter=lambda f: f.GetTypeName() == "SMD", bom=components, correctionFile=correctionpatterns)
    print("storing pos data")
    posDataToFile(posData, os.path.join(outputdir, expandNameTemplate(nametemplate, "pos", loadedBoard) + ".csv"))
    print("collecting solder types")
    types = collectSolderTypes(loadedBoard)
    print("converting bom to Xsv")
    bomToXsv(bom, os.path.join(outputdir, expandNameTemplate(nametemplate, "bom", loadedBoard) + ".csv"), nboards, types)
    bomToXsv(bom, os.path.join(outputdir, expandNameTemplate(nametemplate, "bom", loadedBoard) + ".txt"), nboards, types, delim='\t')
    generateHanwhaSSA(bom, posData, preset, board, os.path.join(outputdir, expandNameTemplate(nametemplate, "hanwha-top", loadedBoard) + ".ssa"), panel_side='T')
    generateHanwhaSSA(bom, posData, preset, board, os.path.join(outputdir, expandNameTemplate(nametemplate, "hanwha-bottom", loadedBoard) + ".ssa"), panel_side='B')
