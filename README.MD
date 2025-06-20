# AE Mask To Premiere

This mask makes it possible to bring an animated mask from Adobe After Effects to Adobe Premiere by injecting a mask into Premiere clipboard data.

# Usage
[Detailed instructions are available here](Instructions/README.MD)

This script currently only works on MacOS, unfortunately. Premiere on Windows seems less willing to modified clipboard data, and I don't have the deep knowledge of Windows' inner workings to find a workaround. If someone out there knows a way to do it, please get in touch.  

# Why?
After Effects and Premiere are very cross-compatible, but masks are an odd exception to that rule. Which is a shame, because AE has superior motion tracking and mask tools. If I need to blur out a logo or whatever, doing it in AE is usually faster. However, to get the mask back into Premiere, you either need to render it out (often a big file) or dynamic link it (reduces performance, makes relinking slightly more annoying). Also, it's nice to have everything in one place.    

# Disclaimer
Use this script at your own risk. It manipulates Premiere data in ways that Adobe didn't intend, and there's always a chance that something could change in a future version of Premiere and break it. If anything bad happens, I am not responsible.