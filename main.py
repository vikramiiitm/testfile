import subprocess
import requests
import cv2
import numpy as np
import win32print
import win32ui
from dbr import *
import time
import pandas as pd
from win32con import DM_ORIENTATION, DMORIENT_LANDSCAPE, SRCCOPY, BLACK_PEN, TRANSPARENT, WHITE_BRUSH
from ctypes import windll, Structure, c_float, byref
from ctypes.wintypes import LONG

''' Pro
1. Get IP of connected phone
2. Check Available Printer
3. Capture Photo stamp
4. Scan the barcode, get ref number
6. Search against ref number in excel
7. Get value from excel and print in stamp
'''

debug=False
def get_device_ip():
    """Gets the IP address of the connected device using ADB."""

    try:
        result = subprocess.run(["adb.exe", "shell", "ifconfig"], capture_output=True)
        print(result)
        output = result.stdout.decode()
        if debug:
            print(f"\n\nadb output: {output} \n\n")

        # Parse the output to extract the IP address
        for line in output.splitlines():
            print(line)
            if "wlan0" in line:
                # Find the next line containing "inet addr"
                for next_line in output.splitlines()[output.splitlines().index(line) + 1:]:
                    if "inet addr:" in next_line or "ccmni2" in next_line:
                        '''with wifi
                            wlan0
                                inet addr:192.168.29.21  Bcast:192.168.29.255  Mask:255.255.255.0 
                                inet addr:192.168.29.21  Bcast:192.168.29.255  Mask:255.255.255.0
                        '''
                        ''' without wifi
                            ccmni2    Link encap:UNSPEC
                                inet addr:100.95.141.196  Mask:255.0.0.0
                                inet6 addr: 2401:4900:55a2:18dc:dc38:4a44:b0af:728b/64 Scope: Global
                                inet6 addr: fe80::dc38:4a44:b0af:728b/64 Scope: Link
                                UP RUNNING NOARP  MTU:1500  Metric:1
                                RX packets:96752 errors:0 dropped:0 overruns:0 frame:0
                                TX packets:29000 errors:0 dropped:0 overruns:0 carrier:0
                                collisions:0 txqueuelen:1000
                                RX bytes:102249064 TX bytes:10128480
                        '''
                        # Extract the IP address
                        ip_address = next_line.split()[1].split(":")[1]
                        return ip_address

                # IP address not found in next line
                print("IP address not found for wlan0 interface")
                return None

        # wlan0 not found
        print("wlan0 interface not found")
        return None

    except Exception as e:
        print("Error getting device IP:", str(e))
        return None
# Replace the below URL with your own. Make sure to add "/shot.jpg" at last.

def capture_photo(ip=None):
    filename="captured_photo.jpg"
    if not ip:
        exit(0)
    url = f"http://{ip}:8080/shot.jpg"
    """Captures a photo from the IP camera and saves it to the specified filename.

    Args:
        filename (str, optional): The filename to save the photo. Defaults to "captured_photo.jpg".

    Returns:
        numpy.ndarray: The captured image as a NumPy array.
    """

    img_resp = requests.get(url)
    img_arr = np.array(bytearray(img_resp.content), dtype=np.uint8)
    img = cv2.imdecode(img_arr, -1)

    # Resize the image if needed
    # img = imutils.resize(img, width=1000, height=1800)

    # Save the photo
    cv2.imwrite(filename, img)
    print(f"Photo saved as {filename}")

    return img

def scan_barcode(filename):
    data = {}
    BarcodeReader.init_license(
        't0068lQAAAJjRSR7zR0cWZB2EpQhYsioGSFbQAlGV5iUjFQ9VBxh8QTU02FpSwicmdcuAnQTAfS5QWAEOBGCP5kdYv2cM6HA=;t0068lQAAADk5uQYPOxgZPZTg3Q/8w9Fw08L+K5ptt41B2WijIOMUleSGJTj5Alch7age51koH2YKKoC66tYhDba+d0336E0=')
    dbr_reader = BarcodeReader()
    filename = "captured_photo.jpg"
    img = cv2.imread(filename)
    try:
        start = time.time()
        dbr_results = dbr_reader.decode_file(filename)
        elapsed_time = time.time() - start

        if dbr_results != None:
            for text_result in dbr_results:
                # print(textResult["BarcodeFormatString"])
                if debug:
                    print('Dynamsoft Barcode Reader: {}. Elapsed time: {}ms'.format(
                        text_result.barcode_text, int(elapsed_time * 1000)))
                data = text_result.barcode_text
                points = text_result.localization_result.localization_points
                cv2.drawContours(
                    img, [np.intp([points[0], points[1], points[2], points[3]])], 0, (0, 255, 0), 2)
                break
            # cv2.imshow('DBR', img)
            return data
        else:
            print("DBR failed to decode {}".format(filename))
    except Exception as err:
        print("DBR failed to decode {}".format(filename))

    return None

def parse_text_to_dict(text):
    import re
    """Parses the given text into a dictionary.

    Args:
        text: The input text string.

    Returns:
        A dictionary containing the parsed key-value pairs.
    """

    result = {}
    lines = text.splitlines()

    for line in lines:
        key_value = line.split(":")
        if len(key_value) == 2:
            key = key_value[0].strip()
            value = key_value[1].strip()
            result[key] = value

    return result

def list_printers():

    try:
        # List all printers with verbose information retrieval
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_CONNECTIONS | win32print.PRINTER_ENUM_LOCAL, None, 1)  # 2 for PRINTER_ENUM_FULL
        #  | win32print.PRINTER_ENUM_LOCAL
        if debug:
            print(f"\n\n Printers Found: {printers} \n\n")
        for index, printer in enumerate(printers):
            printer_name = printer[2]  # Printer name is at index 2
            if debug:
                print(f"Printer #{index+1}: {printer_name}")
            if printer_name == "HP Ink Tank 310 series":
                break
        if printer_name is not None and printer_name:
            return printer_name
        else:
            return None
    except:
        print(f"Error listing printers")
        return None

class XFORM(Structure):
    _fields_ = [("eM11", c_float),
                ("eM12", c_float),
                ("eM21", c_float),
                ("eM22", c_float),
                ("eDx", c_float),
                ("eDy", c_float)]

def rotate_dc_180(hDC, x_center, y_center):
    """Apply 180-degree rotation using transformation matrix."""
    # Set the graphics mode to advanced to support world transformations
    windll.gdi32.SetGraphicsMode(hDC.GetSafeHdc(), 2)  # 2 = GM_ADVANCED
    
    # Create a 180-degree rotation matrix (negative scale on both axes)
    xform = XFORM()
    xform.eM11 = -1.0  # Mirror horizontally
    xform.eM12 = 0.0
    xform.eM21 = 0.0
    xform.eM22 = -1.0  # Mirror vertically
    xform.eDx = 2 * x_center  # Move back to original position
    xform.eDy = 2 * y_center

    # Apply the transformation to the device context
    windll.gdi32.SetWorldTransform(hDC.GetSafeHdc(), byref(xform))

def print_stamp(printer_name, text, x, y, rotate=False):
    try:
        # Get the printer handle
        printer_handle = win32print.OpenPrinter(printer_name)

        # Initialize the printer device context
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)

        # Start the document and the printing page
        hDC.StartDoc('Print Document')
        hDC.StartPage()

        # Set text color (RGB)
        hDC.SetTextColor(0x000000)  # Black color in RGB (0,0,0)

        # Apply 180-degree rotation if specified
        if rotate:
            # Calculate the center of the text (based on where text will be printed)
            text_width, text_height = hDC.GetTextExtent(text)
            x_center = x + text_width // 2
            y_center = y + text_height // 2
            rotate_dc_180(hDC, x_center, y_center)

        # Print the text at the specified position
        hDC.TextOut(x, y, text)

        # End the page and document
        hDC.EndPage()
        hDC.EndDoc()

        # Close the printer handle
        win32print.ClosePrinter(printer_handle)

    except Exception as e:
        print(f"Error printing: {e}")

def load_stamp_data(csv_file_path):
    """Loads the CSV data into a pandas DataFrame with 'stamp_code' as string."""
    return pd.read_csv(csv_file_path, dtype={'stamp_code': str})

def get_stamp_value(stamp_code, df):
    """Fetches the value for a given stamp_code."""
    print(df)
    # Clean up stamp_code and DataFrame values to remove extra quotes and spaces
    df['stamp_code'] = df['stamp_code'].astype(str).str.strip().str.replace('"', '')
    stamp_code = stamp_code.strip().replace('"', '')
    
    # Fetch the row with the matching stamp code
    row = df[df['stamp_code'] == stamp_code]
    print(f"row {row}")
    if not row.empty:
        return row['value'].values[0]
    return None

if __name__ == "__main__":
    ''' Pro
    1. Get IP of connected phone
    2. Check Available Printer
    3. Capture Photo stamp
    4. Scan the barcode, get ref number
    6. Search against ref number in excel
    7. Get value from excel and print in stamp
    '''

    # 1
    ip_address = get_device_ip()
    if ip_address:
        print("Device IP address:", ip_address)
    else:
        print("Error getting device IP")

    #2 Get printer List
    printer_name = list_printers()
    print(printer_name)
    
    # 3 Capture photo and get the captured image
    captured_image = capture_photo(ip_address)
    scan_barcode("captured_photo.jpg")

    #  4. Scan barcode get resul
    scan_result_string = scan_barcode("captured_photo")
    scan_result_dict = parse_text_to_dict(scan_result_string)
    stamp_code = scan_result_dict['E-Stamp Code']
    print(f"stamp Code: {stamp_code}")


    stamp_df = load_stamp_data('data.csv')
    value = get_stamp_value(stamp_code, stamp_df)
    print(f"Value: {value}")

    # Do something with the captured image
    print_stamp(printer_name, stamp_code, 1500, 200)










    


    # Printer print
