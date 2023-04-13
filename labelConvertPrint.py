import win32print
import win32ui
from PIL import Image, ImageWin
import fitz as fz # PyMuPDF
import io
import tkinter as tk
from tkinter import ttk
import tkinterDnD

filterPrinters = False


def pdf_to_img(pdfLocation):
    #This function takes an 8.5"x11" standard shipping label as an input and extracts the raw 4"x6" label from it. (It will still need to be cropped.)
    #Currently only works for UPS
    """
    **Using Fitz
    Open the specified pdf and pull its image data from its first page. 
    Then grab the PDF image-refference number from the first image.
    """
    image = []
    pdf_file = fz.open(pdfLocation)
    for page in pdf_file:
        images = page.get_images()
        for img in images:
            imgRef = img[0]
            print(imgRef)
            """
            **Using Fitz
            Use the image refferenece to extract the image as a byte array.
            """
            base_image = pdf_file.extract_image(imgRef)
            image_bytes = base_image["image"]
            """
            **Using PIL, IO
            Create the Image as a file in memory (using BytesIO to create the file), 
            and using PIL to open it.
            """
            image.append(Image.open(io.BytesIO(image_bytes)))
    return image


def add_margin(pil_img, pad):
    #This function takes in a PIL type image and adds margins on all sides, given the specified pad (in pixels).
    """
    **Using PIL
    Get the image dimmensions, add double the pad to each value.
    and finally paste the original image centered in the new one 
    """
    width, height = pil_img.size
    new_width = 2*pad + width
    new_height = 2*pad + height
    """
    **Using PIL
    Generate a new image using the new dimmensions.
    Paste the original image centered in the new one and return the new padded image.
    """
    result = Image.new(pil_img.mode, (new_width, new_height))
    result.paste(pil_img,(pad,pad))
    return result


def process (file,printer):
    #This function is run whenever a PDF file is loaded into the program. 
    #***Currently doesn't check if it's a PDF!***
    """
    Feed the file location to the pdf_to_img function and store the image that is returned.
    
    Crop the label to 800x1202 (The extra 2 pixels somehow make it look a little better.)

    Note: UPS Labels seem to always contain a label that can be cropped to 800x1200 perfectly;
        As I get labels from more cariers as well as any other variations of UPS labels that exist,
        I can make the cropping function of this a bit more elegant...
    
    Rotate it so it's vertical and add a 4 pixel margin using the add_margin function.
    """

    if file[0] == '{' and file[-1] == '}': #For some reason, tkdnd adds braces around the file location string if it contains spaces. We need to remove those.
            file=file[1:-1] 

    images = pdf_to_img(file)
    left = 0
    top = 0
    right = 800
    bottom = 1202
    for image in images:
        image = image.rotate(-90, Image.NEAREST, expand = 1)
        image = image.crop((left, top, right, bottom))
        image = add_margin(image, 4)
        
        #Take the rotated, cropped, and maginalized image and feed it to the printimg function...
        #Along with the printer name that was passed into THIS function.

        printimg(image,printer)

def save_pref():
    #This function saves user preferences (currently only the last used printer)
    #This function is currently only called when closing the program... 
    #I'll move the root.destroy out of here later.
    data = prints.get()
    file = open("prefs.txt","w")
    file.write(data)
    file.close()
    root.destroy()

def getprinters():
    #This function collects all printers that are set up on this PC and filters out the ones that
    #aren't using 4"x6" paper.

    """
    **Using win32ui, win32print
    Create an array that will be returnned at the end of the function.
    
    Then initialize some constants refering to index values 
    listed in microsofts windows api documentation under "getdevicecaps".
    """
    printOut = []
    HORZRES = 8
    VERTRES = 10
    LOGPIXELSX = 88
    LOGPIXELSY = 90
    """
    Use win32print to Enumerate all printers installed on this PC.

    Iterate throgh them and store the value at position 2 (printer name) of the 
    list for each item into a separate list.
    """
    printers = [printer[2] for printer in win32print.EnumPrinters(2)]
    
    """
    Iterate through the new list of printer names.
    For each name, generate a Device Context and make it a new printer context using the name.
    Using that printer context pull the print area of the printer, as well as the DPI to find
    the physical dimmentions of the page in inches.

    If the printer is a 4"x6" printer, store the name in an array and return the filtered
    array of names when all printers have been checked.
    """
    if filterPrinters == True:
        i = 0
        while (i<len(printers)):
            printer = printers[i]
            #print(printer)
            hDC = win32ui.CreateDC ()
            hDC.CreatePrinterDC (printer)
            prInX = hDC.GetDeviceCaps(HORZRES)/hDC.GetDeviceCaps (LOGPIXELSX)
            prInY = hDC.GetDeviceCaps(VERTRES)/hDC.GetDeviceCaps (LOGPIXELSY)
            if (prInX==4 and prInY==6):
                printOut.append(printer)
            hDC.DeleteDC()
            i = i+1
        return printOut
        #return printers
    else:
        return printers

def printimg(img,printer):
    #This function takes the image to be printed and a printer name as inputs.
    #It scales the image to fit on the page properly and sends it off to the specified printer.

    #Most of this function is from Tim Golden's Single Image example found here: http://timgolden.me.uk/python/win32_how_do_i/print.html
    #I will not document this for now, as Tim has provided his own comments.
    #As I update, I will comment this function in a manner consistent with all others.


    # Constants for GetDeviceCaps
    #
    #
    # HORZRES / VERTRES = printable area
    #
    HORZRES = 8
    VERTRES = 10
    #
    # PHYSICALWIDTH/HEIGHT = total area
    #
    PHYSICALWIDTH = 110
    PHYSICALHEIGHT = 111

    #printer_name = win32print.GetDefaultPrinter ()
    printer_name = printer

    #
    # You can only write a Device-independent bitmap
    #  directly to a Windows device context; therefore
    #  we need (for ease) to use the Python Imaging
    #  Library to manipulate the image.
    #
    # Create a device context from a named printer
    #  and assess the printable size of the paper.
    #
    hDC = win32ui.CreateDC ()
    hDC.CreatePrinterDC (printer_name)
    printable_area = hDC.GetDeviceCaps (HORZRES), hDC.GetDeviceCaps (VERTRES)
    printer_size = hDC.GetDeviceCaps (PHYSICALWIDTH), hDC.GetDeviceCaps (PHYSICALHEIGHT)

    #
    # Open the image, rotate it if it's wider than
    #  it is high, and work out how much to multiply
    #  each pixel by to get it as big as possible on
    #  the page without distorting.
    #
    bmp = img
    if bmp.size[0] > bmp.size[1]:
        bmp = bmp.rotate (90)

    ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
    scale = min (ratios)

    #
    # Start the print job, and draw the bitmap to
    #  the printer device at the scaled size.
    #
    hDC.StartDoc ('label')
    hDC.StartPage ()

    dib = ImageWin.Dib (bmp)
    scaled_width, scaled_height = [int (scale * i) for i in bmp.size]
    x1 = int ((printer_size[0] - scaled_width) / 2)
    y1 = int ((printer_size[1] - scaled_height) / 2)
    x2 = x1 + scaled_width
    y2 = y1 + scaled_height
    dib.draw (hDC.GetHandleOutput (), (x1, y1, x2, y2))

    hDC.EndPage ()
    hDC.EndDoc ()
    hDC.DeleteDC ()

"""
**Using tkinterDnD, tkinter
Setup the Tkinter UI...

As specified in the tkinterDnD documentaion, the initial .tk object must be initialized from
the tkinterDnD library. 
Note: tkiterDnD is the library that provided drag-n-drop support for files.
"""    

root = tkinterDnD.Tk()
root.iconbitmap("myIcon.ico")
root.title("SHIP THIS!")
stringvar = tk.StringVar()
stringvar.set('Drag and drop the PDF here!')

def drop(event):
    # This function runs when a file is dropped in the specified frame.

    process(event.data,prints.get())

    #the following is some bullshit that should be changed...
    #The status bar in the bottom of the window should remain read-only; this is how it is written to.
    status.configure(state='normal')
    status.insert(tk.END,"Sent to be printed...")
    root.after(1000, lambda: status.delete("1.0","end"))
    root.after(1100, lambda: status.configure(state='disabled'))

def startup():
    prs = getprinters()
    status.configure(state="normal")
    status.delete("1.0","end")
    root.after(100, lambda: status.configure(state='disabled'))
    prints['values'] = prs


#initialize the combobox and fill it with printers
n = tk.StringVar()
prints = ttk.Combobox(root, width = 27, textvariable = n)
#prs = getprinters()

#Read the preferences file and use it to write the last-used printer name into the combobox
#prints['values'] =  prs
fi = open("prefs.txt","a")
fi.close()
fi = open("prefs.txt","r")
prints.set(fi.read())
fi.close()
prints.pack()
prints.current()

#Initialize the label that files are drag-n-dropped onto (notice the "ondrop=drop")
#that means that when a file is dropped there, it takes that file location and feeds it to
#the drop function.
label_2 = ttk.Label(root, ondrop=drop, textvar=stringvar, padding=50, relief="solid")
label_2.pack(fill="both", expand=True, padx=10, pady=10)

#Initialize the status window.
status = tk.Text(root,height=1,width=20)
status.pack(fill="both", expand=True, padx=5, pady=5)
status.insert(tk.END,"Finding printers, please wait...")
status.configure(state="disabled")

root.resizable(False,False) #Make the window non-resizable.
root.protocol("WM_DELETE_WINDOW",save_pref) #Register the save_pref function to run on close.
root.after(100,startup)
root.mainloop() #Run the tkinter loop.