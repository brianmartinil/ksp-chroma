# ksp-chroma

Lights up your keyboard to make playing Kerbal Space Program a lot easier. Currently only supports Razer Chroma Keyboards. If you want me to add support for other devices as well, you'll have to send me one. I can send it back after I'm finished implementing the code.

## Features

- Function keys 1 to 0 are only lit, if the underlying action group actually does anything. The keys are displayed in two different colors, depending on whether the action group is toggled or not.
- The keys for SAS, RCS, Gears, Lights and the Brakes are lit up in different colors, indicating if the respective system was activated or not.
- The amount of resources in the current stage is displayed on your keypad and the keys to the left of it (PrtScr, ScrLk, ..., PageDown)
- The color of W, A, S, D, E and Q varies slightly depending on whether you're in precision or normal steering mode
- The keys for timewarp control are lit either red for physics timewarp or green for normal timewarp

Due to the difference between the .NET-Version in the game and the one used by the C# library for the Chroma SDK, I had to create two parts for this mod. The mod itself sends the keyboard color scheme configuration to a small server app, that applies it to the keyboard. 

The mod is still very beta, so let me know if you experience any difficulties when using it.
