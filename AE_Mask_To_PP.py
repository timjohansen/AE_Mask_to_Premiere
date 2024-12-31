import platform
if platform.system() == "Darwin":
    from AppKit import NSPasteboard, NSString, NSStringPboardType   # From pyobjc
# elif platform.system() == "Windows":
#     import win32clipboard as w32cb                                 # From pywin32
else:
    print("Sorry, but this script only works with MacOS. :(")
import xml.etree.ElementTree as ET
import struct
import base64
import json
# import typing

ticks_per_second = 254016000000


class MaskPoint:
    def __init__(self, vertex: tuple[float, float], out_tan: tuple[float, float], in_tan: tuple[float, float]):
        self.vertex = vertex
        self.out_tan = out_tan
        self.in_tan = in_tan


class MaskShape:
    def __init__(self, mask_points: list[MaskPoint], is_closed: bool):
        self.points = mask_points
        self.is_closed = is_closed


class MaskKeyframe:
    def __init__(self, time: float, shape: MaskShape):
        self.time = time
        self.shape = shape


def paste_from_premiere():
    if platform.system() == "Darwin":
        pb = NSPasteboard.generalPasteboard()

        for t in pb.pasteboardItems()[0].types():
            data = pb.pasteboardItems()[0].stringForType_(t)
            try:
                parsed_data = ET.fromstring(data)
                if parsed_data.tag == "PremiereData":
                    premiere_data = pb.pasteboardItems()[0].stringForType_(t)
                    return t, premiere_data
            except ET.ParseError:
                continue
            except TypeError:
                continue
        print("Invalid Premiere clipboard data")
        exit()
    # Windows
    # elif platform.system() == "Windows":
    #     w32cb.OpenClipboard()
    #     premiere_data = w32cb.GetClipboardData(49879)
    #     owner = w32cb.GetClipboardOwner()
    #     w32cb.CloseClipboard()
    #     return owner, premiere_data


def get_start_tick(premiere_clipboard: str):
    parsed_data = ET.fromstring(premiere_clipboard)
    iterator = parsed_data.iterfind("VideoClip")
    while True:
        try:
            node = next(iterator)
            clip_node = node.find("Clip")
            in_point_node = clip_node.find("InPoint")
            if in_point_node is not None:
                return int(in_point_node.text)
        except StopIteration:
            print("Video Clip data not found on clipboard.")
            exit()


def paste_from_ae():
    if platform.system() == "Darwin":
        pb = NSPasteboard.generalPasteboard()
        clipboard_string = pb.stringForType_(NSStringPboardType)
    # elif platform.system() == "Windows":
    #     w32cb.OpenClipboard()
    #     clipboard_string = w32cb.GetClipboardData(w32cb.CF_TEXT).decode()
    #     w32cb.CloseClipboard()

    try:
        pasted_json_data = json.loads(clipboard_string)
    except json.decoder.JSONDecodeError:
        print("Clipboard does not appear to contain AE mask data.")
        return
    processed_keys = []
    for ae_key in pasted_json_data:
        points = []
        for point in ae_key['points']:
            vertex = point['vertex'].split(sep=',')
            vertex = (float(vertex[0]), float(vertex[1]))
            out_tan = point['out'].split(sep=',')
            out_tan = (float(out_tan[0]), float(out_tan[1]))
            in_tan = point['in'].split(sep=',')
            in_tan = (float(in_tan[0]), float(in_tan[1]))
            points.append(MaskPoint(vertex, out_tan, in_tan))
        processed_keys.append(MaskKeyframe(float(ae_key['time']), MaskShape(points, ae_key['closed'])))
    return processed_keys


def mask_shape_to_b64(shape: MaskShape, is_first_key=False):
    # The first few bytes are sometimes b'\x6b\x63\x69\x6e\x02\x00\x00\x00', but the below seems to work consistently.
    raw_data = b'\x32\x63\x69\x6e\x02\x00\x00\x00'
    if shape.is_closed:
        raw_data += b'\x00'
    else:
        raw_data += b'\x02'
    raw_data += struct.pack('>i', len(shape.points))    # '>' specifies big-endian
    raw_data += b'\x00\x00\x00'
    for point in shape.points:
        if point.in_tan == point.vertex and point.out_tan == point.vertex:
            raw_data += b'\x00\x00\x00\x00'
        else:
            raw_data += b'\x01\x00\x00\x00'

        raw_data += struct.pack('f', point.vertex[0])
        raw_data += struct.pack('f', point.vertex[1])
        raw_data += struct.pack('f', point.in_tan[0])
        raw_data += struct.pack('f', point.in_tan[1])
        raw_data += struct.pack('f', point.out_tan[0])
        raw_data += struct.pack('f', point.out_tan[1])
        raw_data += b'\x01\x00\x00\x00'                  # Purpose unknown

    return base64.b64encode(raw_data)


def create_premiere_keyframes(keyframes: list[MaskKeyframe], start_tick) -> str:
    string = ""
    for i in range(len(keyframes)):
        key = keyframes[i]
        string += str(int(key.time * ticks_per_second + start_tick)) + ','
        string += mask_shape_to_b64(key.shape, i == 0).decode() + ';'
    return string


def modify_premiere_clipboard_data(premiere_clipboard: str, encoded_keyframes: str):
    root = ET.fromstring(premiere_clipboard)
    iterator = root.iterfind("ArbVideoComponentParam")
    while True:
        try:
            node = next(iterator)
            keyframes_node = node.find("Keyframes")
            if keyframes_node is not None:
                keyframes_node.text = encoded_keyframes
                break
        except StopIteration:
            print("Keyframes not found in Premiere clipboard")
            exit()

    return str(ET.tostring(root).decode())


def replace_clipboard(keyframe_data: str, premiere_clipboard_data: str, premiere_UTI: str):
    final_string = modify_premiere_clipboard_data(premiere_clipboard_data, keyframe_data)
    if platform.system() == "Darwin":
        allocated_string = NSString.alloc().initWithString_(final_string)
        pasteboard = NSPasteboard.generalPasteboard()
        pasteboard.clearContents()
        pasteboard.declareTypes_owner_([premiere_UTI], None)
        if pasteboard.pasteboardItems()[0].setString_forType_(allocated_string, premiere_UTI):
            return
        else:
            print("Couldn't replace clipboard data.")
    # Windows
    # elif platform.system() == "Windows":
    #     w32cb.OpenClipboard(premiere_UTI)
    #     w32cb.EmptyClipboard()
    #     w32cb.SetClipboardData(49876, final_string)
    #     w32cb.SetClipboardData(49879, final_string)
    #     w32cb.CloseClipboard()


input("In Premiere, add a keyframed mask to the clip you want to have the mask, and copy it to the clipboard.\n\
    Press enter in this window when ready.")
premiere_UTI, premiere_clipboard_data = paste_from_premiere()
input("In After Effects, select the mask you want to copy, run the JSX script. and copy the results. \n\
    Press enter in this window when ready.")
keys = paste_from_ae()
start_tick = get_start_tick(premiere_clipboard_data)
encoded_keys = create_premiere_keyframes(keys, start_tick)
replace_clipboard(encoded_keys, premiere_clipboard_data, premiere_UTI)
print("Mask is ready to paste into Premiere.")