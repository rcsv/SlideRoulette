# SlideRoulette

*SlideRoulette* is a unique PowerPoint VBA tool designed to add an element of surprise to your presentations. By using this tool, you can randomize slide order with a roulette-like effect, making your presentations more engaging and interactive.

## Features

- **Randomize Slides**: Start a roulette effect that randomize the slide order.
- **Unique Presentations**: Each presentation becomes a new experience as slides are shown in different orders.
- **Sound Effects**: Engage your audience further with roulette-start, roulette-stop, and final-sound effects.

## How to use

1. **Setup**: Download the VBA module and integrate it into your PowerPoint presentation file. (Save as *.pptm)
2. **Start Roulette**: Trigger the roulette effect to start randomizing the slides.
3. **Stop Roulette**: Stop the roulette effect to land on a specific slide.
4. **Restart**: Initiate the effect again for another round of surprise.

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
