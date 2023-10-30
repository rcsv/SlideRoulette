# SlideRoulette

*SlideRoulette* is a unique PowerPoint VBA tool designed to add an element of surprise to your presentations. By using this tool, you can randomize slide order with a roulette-like effect, making your presentations more engaging and interactive.

## Features

<img style="float:right;" src="Roulette-Demo.gif" />

- **Randomize Slides**: Start a roulette effect that randomize the slide order.
- **Unique Presentations**: Each presentation becomes a new experience as slides are shown in different orders.
- **Sound Effects**: Engage your audience further with roulette-start, roulette-stop, and final-sound effects.

## How to use SlideRoulette

### 1. **Setup**

* Import the provided `.bas` file into your PowerPoint's VisualBasic for Applications (VBA) environment.
* Ensure that you've set the appropriate path to your sound files within the `soundpath` constant in the imported module.

### 2. Save as Macro-Enabled Presentation

Before integrating SlideRoulette functionalities, make sure to save your PowerPoint presentation in a macro-enabled format. This will ensure that all the VBA functionalities are retained.

- Click on **`File`** in the top left corner.
- Navigate to **`Save As`**.
- From the dropdown, select the location you want to save to.
- In the "Save as type" dropdown, choose **'PowerPoint Macro-Enabled Presentation (*.pptm)'**.
- Click **`Save`**.

### 2. Integrate with Your Presentation
For seamless usage of **SlideRoulette**, you will need to create three buttons within your PowerPoint slides:

#### Start Roulette Button:
Triggers the roulette process, shuffling through slides randomly.

#### Stop Roulette Button:
Halts the roulette process, landing on a slide for the presentation.

#### Reset History Button:
Clears the history of slides that have been landed on by the roulette. This is especially useful if you're planning to run the roulette multiple times within a single presentation sesion.

##### Creating Buttons:

1. **Navigate** to the slide where you'd like to place the buttons.
2. **Go** to **`Insert`** > **`Shapes`** > **`Action Buttons`** and choose a button shape that fits your presentation style.
3. **Draw** the button on the slide.
4. **Assign Macro**: After placing the button, a dialog box should pop up. Assign the corresponding macro (**`StartRoulette`**, **`StopRoulette`**, or **`ResetHistory`**) to the button.
5. **Label**: Optionally, you can right-click on the button to edit text and label it accordingly.

### 3. Run Your Presentation

Once your buttons are in place and the macros are imported, start your slideshow. Use the **Start** and **Stop** buttons as needed during your presentation. If you wish to clear the history and re-run the roulette, simply click the **Reset History** button.

## Requirements
- Microsoft PowerPoint (Version 2019 or later)
- Basic knowledge of VBA (for installation)

## Installation

1. Download the `SlideRoulette.bas` file from this repository.
2. Open your PowerPoint presentation.
3. Press `ALT + F11` to open the VBA editor.
4. Right-Click on `VBAProject (YourPresentationName)` > `Import File` and choose the downloaded `SlideRoulette.bas` file.
5. Close the VBA editor and you're ready to go!

## Limitations

- **Escape Key Behavior**: While in presentation mode, pressing the escape key may cause PowerPoint to crash. It's recommended to avoid using the escape key and instead navigate using the provided controls.
- **PowerPoint Limitations**: SlideRoulette operates within the constraits of PowerPoint's capabilities. Ensure your presentation adheres to standard PowerPoint guidelines to ensure optimal performance.
- **Exhausted Slide Behavior**: After utilizing all slides in the roulette, SlideRoulette might behave erratically or "run wild". It's advisable to reset the roulette or end the presentation mode before reaching this state.

## Contributing

We welcome contributions! If you find any bugs or wish to propose new featurs, please create an issue or submit a pull request.

## License

MIT License. See `LICENSE` for more information.

## Acknowledgements

Special thanks to Rcsvpg for the initial development and to all our contributors for their support.

thank you
