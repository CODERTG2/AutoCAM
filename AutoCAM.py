"""This file acts as the main module for this script."""

import traceback
import adsk.core
import adsk.fusion
import adsk.cam

# Initialize the global variables for the Application and UserInterface objects.
app = adsk.core.Application.get()
ui  = app.userInterface

def run(_context: str):
    """This function is called by Fusion when the script is run."""

    try:
        doc = app.activeDocument
        design = adsk.fusion.Design.cast(doc.products.itemByProductType('DesignProductType'))
        cam = adsk.cam.CAM.cast(doc.products.itemByProductType('CAMProductType'))
        # ui.messageBox(f'"{doc.name}" is the active Document.')

        if not cam:
            ui.messageBox('Switch to the Manufacture workspace or open a file with CAM data.')
            return

        setups = cam.setups
        setupInput = setups.createInput(adsk.cam.OperationTypes.MillingOperation)
        newSetup = setups.add(setupInput)

        newSetup.stockMode = 1

        # get_params(newSetup)

        params = newSetup.parameters
        params.itemByName('job_stockMode').expression = "'default'"
        params.itemByName('job_stockOffsetSides').expression = '0 in'
        params.itemByName('job_stockOffsetTop').expression = '0 in'
        params.itemByName('job_stockOffsetBottom').expression = '0 in'

        ui.messageBox(f"Stock created!")

        params.itemByName('wcs_orientation_mode').expression = "'axesZX'"
        params.itemByName('wcs_orientation_axesZX_unselected_default').expression = "'model'"
        params.itemByName('wcs_orientation_flipZ').value.value = False
        params.itemByName('wcs_origin_mode').expression = "'stockPoint'"
        params.itemByName('wcs_origin_boxPoint').expression = "'bottom 1'"

        ui.messageBox(f"Origin created!")

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
    ui.messageBox('\n'.join(param_names))
