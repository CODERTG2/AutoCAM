"""This file acts as the main module for this script."""

import traceback
import adsk.core
import adsk.fusion
import adsk.cam

# Initialize the global variables for the Application and UserInterface objects.
app = adsk.core.Application.get()
ui  = app.userInterface

COOLANT = "''"
BOTTOM_HEIGHT = ''
DEPTH_PASSES = 0.0
SPINDLE_SPEED = ''
FEED_PLUNGE = ''

def run(_context: str):
    """This function is called by Fusion when the script is run."""

    try:
        doc = app.activeDocument
        cam = adsk.cam.CAM.cast(doc.products.itemByProductType('CAMProductType'))
        design = adsk.fusion.Design.cast(doc.products.itemByProductType('DesignProductType'))

        if not cam:
            ui.messageBox('Switch to the Manufacture workspace or open a file with CAM data.')
            return

        app.log(f"CAM created!")

        setups = cam.setups
        setupInput = setups.createInput(adsk.cam.OperationTypes.MillingOperation)
        newSetup = setups.add(setupInput)

        app.log(f"Setup created!")

        newSetup.stockMode = 1

        app.log(f"Stock mode set!")

        # get_params(newSetup)

        params = newSetup.parameters
        params.itemByName('job_stockMode').expression = "'default'"
        params.itemByName('job_stockOffsetSides').expression = '0 in'
        params.itemByName('job_stockOffsetTop').expression = '0 in'
        params.itemByName('job_stockOffsetBottom').expression = '0 in'

        app.log(f"Stock created!")

        params.itemByName('wcs_orientation_mode').expression = "'axesZX'"
        params.itemByName('wcs_orientation_axesZX_unselected_default').expression = "'model'"
        params.itemByName('wcs_orientation_flipZ').value.value = False
        params.itemByName('wcs_origin_mode').expression = "'stockPoint'"
        params.itemByName('wcs_origin_boxPoint').expression = "'bottom 1'"

        app.log(f"Origin created!")

        # Calculate Stock Z Height from Model Bounding Box
        rootComp = design.rootComponent
        minZ = float('inf')
        maxZ = float('-inf')
        
        # Check if there are bodies
        if rootComp.bRepBodies.count == 0:
             ui.messageBox('No bodies found in root component.')
             return

        for body in rootComp.bRepBodies:
            bbox = body.boundingBox
            if bbox.minPoint.z < minZ: minZ = bbox.minPoint.z
            if bbox.maxPoint.z > maxZ: maxZ = bbox.maxPoint.z
        
        stock_z_height_cm = maxZ - minZ
        
        # User defined settings
        (material, cancelled) = ui.inputBox("Is this polycarb or aluminum?", "Material", "polycarb")
        if cancelled: return

        (passes, cancelled) = ui.inputBox("How many passes?", "Passes", "4")
        if cancelled: return

        COOLANT = "'disabled'"
        BOTTOM_HEIGHT = '-0.01 in'
        DEPTH_PASSES = f"{stock_z_height_cm / int(passes) * 0.393701} in"
        if material.lower() == "polycarb":
            SPINDLE_SPEED = '16000 rpm'
            FEED_PLUNGE = '8.33 in/min'
        elif material.lower() == "aluminum":
            SPINDLE_SPEED = '20000 rpm'
            FEED_PLUNGE = '6 in/min'
        else:
            ui.messageBox('Invalid material selected.')
            return
        
        config = {
            'SPINDLE_SPEED': SPINDLE_SPEED,
            'FEED_PLUNGE': FEED_PLUNGE,
            'COOLANT': COOLANT,
            'BOTTOM_HEIGHT': BOTTOM_HEIGHT,
            'DEPTH_PASSES': DEPTH_PASSES
        }

        bore(doc, newSetup, cam, config)

        app.log(f"Bore created!")

        manual_nc_input = newSetup.operations.createInput('manual')
        manual_nc_op = newSetup.operations.add(manual_nc_input)
            
        param = manual_nc_op.parameters.itemByName('manualType')
        if param:
            param.expression = "'stop'"
            app.log("Manual NC Stop created successfully!")
        else:
            app.log("Error: 'manualType' parameter not found.")

        pocket(doc, newSetup, cam, config)

        app.log(f"Pocket created!")

        # Second Manual NC Stop (between pocket and contour)
        manual_nc_input2 = newSetup.operations.createInput('manual')
        manual_nc_op2 = newSetup.operations.add(manual_nc_input2)
        param2 = manual_nc_op2.parameters.itemByName('manualType')
        if param2:
            param2.expression = "'stop'"
            app.log("Manual NC Stop 2 created successfully!")
        
        contour(doc, newSetup, cam, config)

    except:  #pylint:disable=bare-except
        # Write the error message to the TEXT COMMANDS window.
        app.log(f'Failed:\n{traceback.format_exc()}')


def get_params(setup):
    operation = setup.operations.item(0)
    parameters = operation.parameters
        
    param_names = []
    for i in range(parameters.count):
        param = parameters.item(i)
        param_names.append(param.name)

    param_names.sort()
    app.log('\n'.join(param_names))

def bore(doc, setup, cam, config):
    """Detect holes in the active design for boring operations."""
    design = adsk.fusion.Design.cast(doc.products.itemByProductType('DesignProductType'))

    if not design:
        ui.messageBox('No active design found.')
        return
        
    rootComp = design.rootComponent
    body = rootComp.bRepBodies.item(0)
    
    planar_faces = [f for f in body.faces if f.geometry.surfaceType == adsk.core.SurfaceTypes.PlaneSurfaceType]
    front_face = max(planar_faces, key=lambda f: f.area)

    circles = []
    for edge in front_face.edges:
        if edge.geometry.curveType == adsk.core.Curve3DTypes.Circle3DCurveType:
            circle = adsk.core.Circle3D.cast(edge.geometry)
            circles.append({
                'edge': edge,
                'center': circle.center,
                'radius': circle.radius,
                'diameter': circle.radius * 2,
                'normal': circle.normal
            })
    
    app.log(f"Found {len(circles)} circular edges.")

    hole_faces = []
    for circle in circles:
        edge = circle['edge']

        for face in edge.faces:
            if face.geometry.surfaceType == adsk.core.SurfaceTypes.CylinderSurfaceType:
                cylinder = adsk.core.Cylinder.cast(face.geometry)
                if abs(cylinder.radius - circle['radius']) < 0.001:
                    hole_faces.append(face)
                    break

    app.log(f"Found {len(hole_faces)} holes.")

    operations = setup.operations
    boring_input = operations.createInput('bore')
    boring_op = operations.add(boring_input)
    
    params = boring_op.parameters

    tool = get_tool(cam)
    if tool:
        # Use the tool in your operation
        boring_op.tool = tool
        app.log("Tool assigned to operation!")
    else:
        app.log("Warning: Could not find the specific tool, using default parameters")

        # Set default tool parameters first
        params.itemByName('tool_diameter').expression = '3.175 mm'
        params.itemByName('tool_type').expression = "'flat end mill'"
        params.itemByName('tool_numberOfFlutes').expression = '2'

    # Set other machining parameters
    params.itemByName('tool_spindleSpeed').expression = config['SPINDLE_SPEED']
    params.itemByName('tool_feedPlunge').expression = config['FEED_PLUNGE']
    params.itemByName('tool_coolant').expression = config['COOLANT']
    params.itemByName('bottomHeight_offset').expression = config['BOTTOM_HEIGHT']
    params.itemByName('strategy').expression = "'bore'"

    params.itemByName('circularFaces').value.value = hole_faces

    cam.generateToolpath(boring_op)
    app.log("Toolpath generated!")

def pocket(doc, setup, cam, config):
    """Create a 2D pocket operation for non-circular through-features."""
    design = adsk.fusion.Design.cast(doc.products.itemByProductType('DesignProductType'))
    if not design:
        ui.messageBox('No active design found.')
        return

    root = design.rootComponent

    # Find the bottom face (largest downward-facing planar face)
    bottom_face = None
    max_area = 0
    for body in root.bRepBodies:
        for face in body.faces:
            if face.geometry.surfaceType == adsk.core.SurfaceTypes.PlaneSurfaceType:
                success, normal = face.evaluator.getNormalAtPoint(face.pointOnFace)
                if success and normal.z < -0.9 and face.area > max_area:
                    max_area = face.area
                    bottom_face = face

    if not bottom_face:
        app.log("No bottom face found, skipping pocket.")
        return

    # Find non-circular inner loops on the bottom face
    # Inner loops = holes/cutouts; circular ones are handled by bore
    pocket_loops = []
    for loop in bottom_face.loops:
        if loop.isOuter:
            continue  # Skip the outer boundary (handled by contour)

        # Check if this loop is purely circular (handled by bore)
        edges = loop.edges
        is_circle = False
        if edges.count <= 2:
            all_arcs = True
            for edge in edges:
                ct = edge.geometry.curveType
                if ct != adsk.core.Curve3DTypes.Circle3DCurveType and ct != adsk.core.Curve3DTypes.Arc3DCurveType:
                    all_arcs = False
                    break
            is_circle = all_arcs

        if not is_circle:
            pocket_loops.append(loop)

    if not pocket_loops:
        app.log("No non-circular pocket features found, skipping pocket.")
        return

    app.log(f"Found {len(pocket_loops)} non-circular pocket loop(s).")

    # Create 2D contour operation for non-circular through-features
    operations = setup.operations
    contour_input = operations.createInput('contour2d')
    contour_op = operations.add(contour_input)

    params = contour_op.parameters

    # Assign tool
    tool = get_tool(cam)
    if tool:
        contour_op.tool = tool
        app.log("Tool assigned to pocket contour operation!")
    else:
        app.log("Warning: Could not find the specific tool, using default parameters")
        params.itemByName('tool_diameter').expression = '3.175 mm'
        params.itemByName('tool_type').expression = "'flat end mill'"
        params.itemByName('tool_numberOfFlutes').expression = '2'

    # Set contour geometry — chain selections for each non-circular loop
    geom_param = params.itemByName('contours')
    cad_contours = adsk.cam.CadContours2dParameterValue.cast(geom_param.value)
    selections = cad_contours.getCurveSelections()

    for loop in pocket_loops:
        chain_sel = selections.createNewChainSelection()
        edge_list = [edge for edge in loop.edges]
        chain_sel.inputGeometry = edge_list

    cad_contours.applyCurveSelections(selections)

    # Set machining parameters
    params.itemByName('tool_spindleSpeed').expression = config['SPINDLE_SPEED']
    params.itemByName('tool_feedPlunge').expression = config['FEED_PLUNGE']
    params.itemByName('tool_coolant').expression = config['COOLANT']
    params.itemByName('bottomHeight_offset').expression = config['BOTTOM_HEIGHT']

    # Multiple depths
    p_do_multiple = params.itemByName('doMultipleDepths')
    if p_do_multiple:
        p_do_multiple.value.value = True

    p_stepdown = params.itemByName('maximumStepdown')
    if p_stepdown:
        p_stepdown.expression = config['DEPTH_PASSES']

    cam.generateToolpath(contour_op)
    app.log("Pocket Contour Toolpath generated!")

def contour(doc, setup, cam, config):
    """Create a 2D contour operation for the outer profile."""
    operations = setup.operations
    contour_input = operations.createInput('contour2d')
    contour_op = operations.add(contour_input)

    params = contour_op.parameters

    tool = get_tool(cam)
    if tool:
        contour_op.tool = tool
        app.log("Tool assigned to contour operation!")
    else:
        app.log("Warning: Could not find the specific tool, using default parameters")
        params.itemByName('tool_diameter').expression = '3.175 mm'
        params.itemByName('tool_type').expression = "'flat end mill'"
        params.itemByName('tool_numberOfFlutes').expression = '2'

    design = adsk.fusion.Design.cast(doc.products.itemByProductType('DesignProductType'))
    if not design:
        ui.messageBox('No active design found.')
        return

    root = design.rootComponent

    # Find the bottom face using LARGEST AREA (not min Z)
    # This ensures pockets don't interfere — pocket floors are smaller
    bottom_face = None
    max_area = 0

    for body in root.bRepBodies:
        for face in body.faces:
            if face.geometry.surfaceType == adsk.core.SurfaceTypes.PlaneSurfaceType:
                success, normal = face.evaluator.getNormalAtPoint(face.pointOnFace)
                if success and normal.z < -0.9:
                    if face.area > max_area:
                        max_area = face.area
                        bottom_face = face

    if bottom_face:
        # Use FaceContourSelection instead of ChainSelection
        # This is simpler and more robust — no manual edge gathering needed
        geom_param = params.itemByName('contours')
        cad_contours = adsk.cam.CadContours2dParameterValue.cast(geom_param.value)
        selections = cad_contours.getCurveSelections()

        face_sel = selections.createNewFaceContourSelection()
        face_sel.inputGeometry = [bottom_face]
        face_sel.loopType = adsk.cam.LoopTypes.OnlyOutsideLoops
        face_sel.sideType = adsk.cam.SideTypes.AlwaysOutsideSideType

        cad_contours.applyCurveSelections(selections)
        app.log(f"Bottom face selected for contour (area: {bottom_face.area:.4f})")
    else:
        app.log("Warning: No bottom face found for contour!")

    # Set machining parameters
    params.itemByName('tool_spindleSpeed').expression = config['SPINDLE_SPEED']
    params.itemByName('tool_feedPlunge').expression = config['FEED_PLUNGE']
    params.itemByName('tool_coolant').expression = config['COOLANT']

    # Bottom Height
    params.itemByName('bottomHeight_offset').expression = config['BOTTOM_HEIGHT']

    # Maximum Roughing Stepdown
    p_do_multiple = params.itemByName('doMultipleDepths')
    if p_do_multiple:
        p_do_multiple.value.value = True

    p_stepdown = params.itemByName('maximumStepdown')
    if p_stepdown:
        p_stepdown.expression = config['DEPTH_PASSES']

    cam.generateToolpath(contour_op)
    app.log("Contour Toolpath generated!")

def get_tool(cam):
    toolLib: adsk.cam.DocumentToolLibrary = cam.documentToolLibrary
    # assuming that there is only 1 tool in the library
    return toolLib.item(0)
