# shippingLabelConvertPrint
Python tool used to extract, crop, and print 4"x6" shipping labels (for use with a thermal printer) from 8.5"x11" labels that are typically provided.


# General Info
The code is was thrown together and is pretty rough at the moment but I've commented throughout to relay the general workflow in its current state.

Currently this (may) only work for UPS labels. I'll verify or add support for FedEx next time I ship something through them.

#Requirements
myIcon.ico must be in the same folder as the python script or the .exe file.

If you are running the python script, please read the dependancyHowTo.txt for a list of dependancies and how to install them.

#How to Use
1. Run the program.

<img src="https://user-images.githubusercontent.com/55807922/231681726-4b010c5e-ab92-42fc-b119-20359b373235.png" width = "250">

2. Select a printer from the dropdown menu.

<img src="https://user-images.githubusercontent.com/55807922/231683436-841577b1-a315-4084-80f7-a5b3faa0150c.png" width = "250">

3. From your file explorer, drag and drop the PDF containing the shipping label that you would like to print.

<img src="https://user-images.githubusercontent.com/55807922/231684157-1de5b90e-bce4-483b-85c3-1ed36fb98b2a.png" width = "250">

The shipping label will then be printed by the designated printer.


