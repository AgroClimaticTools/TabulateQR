import time
from typing import Union

from cv2 import VideoCapture, destroyAllWindows, imshow, waitKey
from pandas import DataFrame, concat
from pyzbar.pyzbar import decode


def scan_qr_code():
    # Open the default camera (you can change the argument to use a different camera)
    cap = VideoCapture(0)
    start_time = time.time()
    
    while True:
        # Capture a frame
        ret, frame = cap.read()

        # Decode QR codes
        decoded_objects = decode(frame)

        # Display the frame
        imshow('QR Code Scanner', frame)
        waitKey(1)

        # Check for QR codes
        if decoded_objects is not None:
            for obj in decoded_objects:
                qrcode = obj.data.decode('utf-8')
                # Release the camera and close the window
                cap.release()
                destroyAllWindows()
                return qrcode
        
        # time out break
        if time.time() - start_time > 5:
            break
    # Release the camera and close the window
    cap.release()
    destroyAllWindows()

def isQRCode(qrCode):
        return len(qrCode) >= 10 and qrCode.isalnum()


def isfloat(value: Union[float, str, int]) -> bool:
    '''
    Check if the value is float
    
    :param value: str/float/int to check for float

    :return bool: if float returns True else False
    '''
    try:
        float(value)
        return True
    except ValueError:
        return False

def convert2StrIntFloat(
        value: Union[float, str, int]) -> Union[float, str, int]:
    if isfloat(value):
        if float(value).is_integer():
            return int(float(value))
        else:
            return float(value)
    else:
        return value

def decodeQRCode(qrCodes: list[str], decode_df: DataFrame):
    qrCode_sep_pos = {}
    decode_dict = decode_df.to_dict()['Decoding']
    for key, _ in decode_dict.items():
        qrCode_sep_pos[key] = qrCodes[0].index(key)
    qrCode_ordered_codes = ['QR Code']+[decode_dict[key] for key, _ in sorted(
        qrCode_sep_pos.items(), key=lambda x: x[1])]
    df_qrCode = []
    for qrCode in qrCodes:
        qrCode_string = qrCode
        data = {'QR Code': qrCode}
        for key, _ in sorted(
            qrCode_sep_pos.items(), key=lambda x: x[1], reverse=True):
            data[decode_dict[key]] = \
                [convert2StrIntFloat(qrCode_string.split(key)[-1])]
            # data[decode_dict[key]] = \
            #     [convert2StrIntFloat(key+qrCode_string.split(key)[-1])]
            qrCode_string = qrCode_string.split(key)[0]
        df_qrCode.append(DataFrame.from_records(data))
    df = concat(df_qrCode).reset_index()
    return (df.loc[:,qrCode_ordered_codes], qrCode_ordered_codes)

'______________________________________________________________________________'

if __name__ == "__main__":
    print(scan_qr_code())
