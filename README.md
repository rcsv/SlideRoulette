# SlideRoulette

*SlideRoulette* is a unique PowerPoint VBA tool designed to add an element of surprise to your presentations. By using this tool, you can randomize slide order with a roulette-like effect, making your presentations more engaging and interactive.

## Features

<img style="float:right;" src="Roulette-Demo.gif" />

- **Randomize Slides**: Start a roulette effect that randomizes the slide order.
- **Unique Presentations**: Each presentation becomes a new experience as slides are shown in different orders.
- **Sound Effects**: Engage your audience further with roulette-start, roulette-stop, and final-sound effects.
- **History Tracking**: SlideRoulette remembers which slides have come up during the roulette, ensuring each slide is only selected once per session.
- **History Display**: Visually track the order in which slides have been presented with the on-slide display feature, allowing both the presenter and audience to see which slides have already been shown.

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

### 3. Integrate with Your Presentation
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

### 4. Run Your Presentation

Once your buttons are in place and the macros are imported, start your slideshow. Use the **Start** and **Stop** buttons as needed during your presentation. If a slide has been landed on, it will be tracked and excluded from future selections until the **Reset History** button is clicked. This feature ensures every slide gets a chance to be displayed without repeats.

If you wish to view which slides have already been displayed, or if your audience has missed some parts of the presentation, you can use the history display feature. This feature updates a text box on the second slide with the numbers of the slides that have been shown. It serves as a quick reference for both presenter and audience to keep track of the covered content.

To clear the history and re-run the roulette, simply click the **Reset History** button. This will erase the memory of previous selections and allow all slides to be available for random selection again.

## History Feature Details

**Tracking and Displaying History**: SlideRoulette includes a comprehensive history tracking and display system. This ensures that once a slide has been selected by the roulette, it won't be chosen again during the same session. This feature enhances the randomness by maintaining the thrill of the roulette without repetition. A text box on the second slide automatically updates to show all the selected slide numbers, creating a visual history for the presenter and the audience.

**Resetting History**: When the roulette is run multiple times during a session, or when starting a new session, the history can be reset. This clears the previously selected slide numbers, making all slides available for selection once again. The corresponding button can be quickly accessed during the presentation to reset the slide selection process.

For detailed steps on how to utilize these new features, please refer to the updated sections 'Integrate with Your Presentation' and 'Run Your Presentation'.

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

- Special thanks to [Rcsvpg](https://github.com/rcsv) for the initial development and to all our contributors for their support.

Thank you.
