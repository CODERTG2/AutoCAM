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

# TODO: select tool, select contour, select holes

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

        COOLANT = "'disabled'"
        BOTTOM_HEIGHT = '-0.01 in'
        DEPTH_PASSES = f"{stock_z_height_cm / 5.0 * 0.393701} in"
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

    cam.generateToolpath(boring_op)
    app.log("Toolpath generated!")
    
    # Show completion message with instructions for changing tool
    show_completion_message()

def show_completion_message():
    """
    Shows a single message when the operation is complete.
    """
    app = adsk.core.Application.get()
    ui = app.userInterface
    
    ui.messageBox(
        'Boring operation created!\n\n' +
        'To change the tool:\n' +
        '• Double-click the operation in the Browser\n' +
        '• Click the "Tool" tab and select your tool\n' +
        '• Click "Generate" to regenerate the toolpath',
        'AutoCAM Complete',
        adsk.core.MessageBoxButtonTypes.OKButtonType,
        adsk.core.MessageBoxIconTypes.InformationIconType
    )

def contour(doc, setup, cam, config):
    """Create a 2D contour operation."""
    operations = setup.operations
    contour_input = operations.createInput('contour2d')
    contour_op = operations.add(contour_input)
    
    params = contour_op.parameters

    # Set tool parameters (same as bore)
    params.itemByName('tool_diameter').expression = '3.175 mm'
    params.itemByName('tool_type').expression = "'flat end mill'"
    params.itemByName('tool_numberOfFlutes').expression = '2'

    # Set machining parameters
    params.itemByName('tool_spindleSpeed').expression = config['SPINDLE_SPEED']
    params.itemByName('tool_feedPlunge').expression = config['FEED_PLUNGE']
    params.itemByName('tool_coolant').expression = config['COOLANT']
    
    # Bottom Height
    params.itemByName('bottomHeight_offset').expression = config['BOTTOM_HEIGHT']
    
    # Maximum Roughing Stepdown
    try:
        # Enable multiple depths
        p_multiple_depths = params.itemByName('operation_strategy_multipleDepth')
        if not p_multiple_depths:
             p_multiple_depths = params.itemByName('multipleDepthsGroup') # Sometimes a group
        
        if params.itemByName('multipleDepthsEnabled'):
            params.itemByName('multipleDepthsEnabled').value.value = True
        
        params.itemByName('maximumStepdown').expression = config['DEPTH_PASSES']
    except:
        app.log('Could not set multiple depths parameters completely.')

    cam.generateToolpath(contour_op)
    app.log("Contour Toolpath generated!")
