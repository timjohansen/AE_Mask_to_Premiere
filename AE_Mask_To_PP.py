import platform
if platform.system() == "Darwin":
    from AppKit import NSPasteboard, NSString, NSStringPboardType   # From pyobjc
# elif platform.system() == "Windows":
#     import win32clipboard as w32cb                                 # From pywin32
else:
    print("Sorry, but this script only works with MacOS. :(")
    exit()
import xml.etree.ElementTree as ET
import struct
import base64
import json
import tkinter
from tkinter import ttk, messagebox, N, S, E, W

time_units_per_second = 254016000000


class MaskPoint:
    def __init__(self,
                 vertex: tuple[float, float],
                 out_tan: tuple[float, float],
                 in_tan: tuple[float, float]):
        self.vertex = vertex
        self.out_tan = out_tan
        self.in_tan = in_tan


class MaskShape:
    def __init__(self,
                 mask_points: list[MaskPoint],
                 is_closed: bool):
        self.points = mask_points
        self.is_closed = is_closed


class MaskKeyframe:
    def __init__(self,
                 time: float,
                 shape: MaskShape):
        self.time = time
        self.shape = shape


def paste_from_premiere():
    if platform.system() == "Darwin":
        pb = NSPasteboard.generalPasteboard()
        pb_type = None

        parsed_data = None
        premiere_data_string = None
        clip_name = ""

        for t in pb.pasteboardItems()[0].types():
            raw_data = pb.pasteboardItems()[0].stringForType_(t)
            try:
                parsed_data = ET.fromstring(raw_data)
                if parsed_data.tag == "PremiereData":
                    premiere_data_string = pb.pasteboardItems()[0].stringForType_(t)
                    pb_type = t
                    break
            except ET.ParseError:
                continue
            except TypeError:
                continue
        if premiere_data_string is None:
            messagebox.showinfo("Error", "Premiere data was not found on clipboard.")
            return None

        iterator = parsed_data.iterfind("ArbVideoComponentParam")

        while True:
            try:
                node = next(iterator)
                name_node = node.find("Name")
                if name_node is None or name_node.text != "Mask Path":
                    continue
                keyframes_node = node.find("Keyframes")
                if keyframes_node is not None:
                    break
            except StopIteration:
                messagebox.showinfo("Error", "No keyframes were found in the clipboard data.")
                return None

        try:
            object_id = int(node.attrib["ObjectID"])
        except:
            messagebox.showinfo("Error", "Could not parse ArbVideoComponentParam ID.")
            return None

        iterator = parsed_data.iterfind("VideoFilterComponent")
        found = False
        while True:
            try:
                node = next(iterator)
                vfc_node = node
                node = node.find("Component")
                node = node.find("Params")
                param_iter = node.iterfind("Param")
                while True:
                    try:
                        param_node = next(param_iter)
                        if int(param_node.attrib["ObjectRef"]) == object_id:
                            object_id = int(vfc_node.attrib["ObjectID"])
                            found = True
                            break
                    except StopIteration:
                        break
            except StopIteration:
                break
        if not found:
            messagebox.showinfo("Error", "Could not find Param node in clipboard data.")
            return None

        iterator = parsed_data.iterfind("VideoFilterComponent")
        found = False
        while True:
            try:
                node = next(iterator)
                vfc_node = node
                node = node.find("SubComponents")
                if node is None:
                    continue
                sub_component_iter = node.iterfind("SubComponent")
                while True:
                    try:
                        subcom_node = next(sub_component_iter)
                        if int(subcom_node.attrib["ObjectRef"]) == object_id:
                            object_id = int(vfc_node.attrib["ObjectID"])
                            found = True
                            break
                    except StopIteration:
                        break
            except StopIteration:
                break
        if not found:
            messagebox.showinfo("Error", "Could not find SubComponent node in clipboard data.")
            return None

        iterator = parsed_data.iterfind("VideoComponentChain")
        found = False
        while True:
            try:
                node = next(iterator)
                vcc_node = node
                node = node.find("ComponentChain")
                node = node.find("Components")
                component_iter = node.iterfind("Component")
                while True:
                    try:
                        com_node = next(component_iter)
                        if int(com_node.attrib["ObjectRef"]) == object_id:
                            object_id = int(vcc_node.attrib["ObjectID"])
                            found = True
                            break
                    except StopIteration:
                        break
            except StopIteration:
                break
        if not found:
            messagebox.showinfo("Error", "Could not find Component node in clipboard data.")
            return None

        iterator = parsed_data.iterfind("VideoClipTrackItem")
        found = False
        while True:
            try:
                node = next(iterator)
                try:
                    ct_node = node.find("ClipTrackItem")
                    node = ct_node.find("ComponentOwner")
                    node = node.find("Components")
                    if object_id == int(node.attrib["ObjectRef"]):
                        found = True
                        node = ct_node.find("SubClip")
                        object_id = int(node.attrib["ObjectRef"])
                except:
                    continue

            except StopIteration:
                break
        if not found:
            messagebox.showinfo("Error", "Could not find Component node in clipboard data.")
            return None

        iterator = parsed_data.iterfind("SubClip")
        found = False
        while True:
            try:
                node = next(iterator)
                clip_node = node.find("Clip")
                object_id = int(clip_node.attrib["ObjectRef"])
                node = node.find("Name")
                clip_name = node.text
                found = True
                break
            except StopIteration:
                break
        if not found:
            messagebox.showinfo("Error", "Could not find SubClip node in clipboard data.")
            return None

        iterator = parsed_data.iterfind("VideoClip")
        found = False
        while True:
            try:
                node = next(iterator)
                if int(node.attrib["ObjectID"]) != object_id:
                    continue
                node = node.find("Clip")
                in_point_node = node.find("InPoint")
                if in_point_node is not None:
                    try:
                        in_point = int(in_point_node.text)
                        found = True
                        break
                    except:
                        messagebox.showinfo("Error", "Could not parse InPoint node.")
                        return None
            except StopIteration:
                break
        if not found:
            messagebox.showinfo("Error", "Could not find VideoClip node in clipboard data.")
            return None

        return premiere_data_string, pb_type, in_point, clip_name

    # elif platform.system() == "Windows":
    #     w32cb.OpenClipboard()
    #     owner = w32cb.GetClipboardOwner()
    #     premiere_data = w32cb.GetClipboardData(49879)
    #     w32cb.CloseClipboard()
    #     return owner, premiere_data


def paste_from_ae():
    if platform.system() == "Darwin":
        pb = NSPasteboard.generalPasteboard()
        clipboard_string = pb.stringForType_(NSStringPboardType)
    # elif platform.system() == "Windows":
    #     w32cb.OpenClipboard()
    #     clipboard_string = w32cb.GetClipboardData(w32cb.CF_TEXT).decode()
    #     w32cb.CloseClipboard()

    try:
        json_object = json.loads(clipboard_string)
    except json.decoder.JSONDecodeError:
        messagebox.showinfo("Error", "Clipboard data does not appear to contain After Effects data.\n"
                                     "This script can only accept data generated by Mask_To_JSON.jsx.")
        return None
    processed_keys = []
    for ae_key in json_object:
        points = []
        for point in ae_key['points']:
            try:
                vertex = point['vertex'].split(sep=',')
                vertex = (float(vertex[0]), float(vertex[1]))
                out_tan = point['out'].split(sep=',')
                out_tan = (float(out_tan[0]), float(out_tan[1]))
                in_tan = point['in'].split(sep=',')
                in_tan = (float(in_tan[0]), float(in_tan[1]))
                points.append(MaskPoint(vertex, out_tan, in_tan))
            except KeyError:
                messagebox.showinfo("Error", "Clipboard data is corrupted or incomplete.")
                return None
        processed_keys.append(MaskKeyframe(float(ae_key['time']), MaskShape(points, ae_key['closed'])))
    return processed_keys


def mask_shape_to_b64(shape: MaskShape, is_first_key=False):
    if shape is None:
        return None

    # The first few bytes are sometimes b'\x6b\x63\x69\x6e\x02\x00\x00\x00', and I'm not sure why,
    # but the sequence below seems to work consistently.
    raw_data = b'\x32\x63\x69\x6e\x02\x00\x00\x00'
    if shape.is_closed:
        raw_data += b'\x00'
    else:
        raw_data += b'\x02'
    raw_data += struct.pack('>i', len(shape.points))    # '>i' specifies big-endian integer.
    raw_data += b'\x00\x00\x00'
    for point in shape.points:
        if point.in_tan == point.vertex and point.out_tan == point.vertex:
            raw_data += b'\x00\x00\x00\x00'
        else:
            raw_data += b'\x01\x00\x00\x00'             # Specifies that the point is part of a curve.

        raw_data += struct.pack('f', point.vertex[0])
        raw_data += struct.pack('f', point.vertex[1])
        raw_data += struct.pack('f', point.in_tan[0])
        raw_data += struct.pack('f', point.in_tan[1])
        raw_data += struct.pack('f', point.out_tan[0])
        raw_data += struct.pack('f', point.out_tan[1])
        raw_data += b'\x01\x00\x00\x00'                  # Purpose unknown

    return base64.b64encode(raw_data)


def create_premiere_keyframes(keyframes: list[MaskKeyframe],
                              start_time) -> str:
    keyframe_string = ""
    for i in range(len(keyframes)):
        key = keyframes[i]
        keyframe_string += str(int(key.time * time_units_per_second + start_time)) + ','
        keyframe_string += mask_shape_to_b64(key.shape, i == 0).decode() + ';'
    return keyframe_string


def modify_premiere_clipboard_data(premiere_clipboard: str,
                                   encoded_keyframes: str):
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
            messagebox.showinfo("Error", "Keyframes were not found in the clipboard data.")
            return None
        except Exception as err:
            messagebox.showinfo("Error", f"An unexpected error occurred: {err}")
            return None

    return str(ET.tostring(root).decode())


def replace_clipboard(keyframe_data: str,
                      premiere_clipboard_data: str,
                      premiere_ID: str):
    final_string = modify_premiere_clipboard_data(premiere_clipboard_data, keyframe_data)
    if final_string is None:
        return
    if platform.system() == "Darwin":
        allocated_string = NSString.alloc().initWithString_(final_string)
        pasteboard = NSPasteboard.generalPasteboard()
        pasteboard.clearContents()
        pasteboard.declareTypes_owner_([premiere_ID], None)
        if pasteboard.pasteboardItems()[0].setString_forType_(allocated_string, premiere_ID):
            return
        else:
            messagebox.showinfo("Error", "Unable to replace clipboard data.")
            return
    # elif platform.system() == "Windows":
    #     w32cb.OpenClipboard(premiere_ID)
    #     w32cb.EmptyClipboard()
    #     w32cb.SetClipboardData(49876, final_string)
    #     w32cb.SetClipboardData(49879, final_string)
    #     w32cb.CloseClipboard()


def premiere_button_clicked():
    global premiere_status
    global premiere_ID
    global premiere_clipboard_data
    global premiere_in_point
    data = paste_from_premiere()
    if data is None:
        return
    premiere_clipboard_data, premiere_ID, premiere_in_point, clip_name = data
    premiere_status.set("Video Clip:\n" + clip_name)
    check_if_ready()


def check_if_ready():
    if ae_clipboard_data is not None and premiere_clipboard_data is not None:
        process_button.config(state=tkinter.NORMAL, text="Click to Copy")
    else:
        process_button.config(state=tkinter.DISABLED, text="Not Ready")


def ae_button_clicked():
    global ae_clipboard_data
    global ae_status
    data = paste_from_ae()
    if data is not None:
        ae_clipboard_data = data
        ae_status.set("Keyframes: " + str(len(ae_clipboard_data)))
        check_if_ready()


def convert_button_clicked():
    if ae_clipboard_data is None or premiere_clipboard_data is None:
        print("This button should not have been clickable")
        return
    encoded_keys = create_premiere_keyframes(ae_clipboard_data, premiere_in_point)
    replace_clipboard(encoded_keys, premiere_clipboard_data, premiere_ID)
    messagebox.showinfo("Success", "Clipboard data was successfully replaced. The mask can now be pasted into Premiere.")


# Globals
premiere_clipboard_data = None
ae_clipboard_data = None
premiere_ID = None
premiere_in_point = None

# UI
window = tkinter.Tk()
window.title("AE Mask to Premiere")
main_frame = ttk.Frame(window, padding=10)
main_frame.grid(column=0, row=0, sticky=(N,W,E,S))

paste_frame = ttk.Frame(main_frame, padding=10)
paste_frame.grid(column=0, row=0, sticky=(N,W,E,S))

premiere_paste_frame = ttk.Frame(paste_frame, padding=5, borderwidth=1, relief="solid")
premiere_paste_frame.grid(column=0, row=0, sticky=(N,S))
premiere_paste_button = ttk.Button(premiere_paste_frame, text="Paste from Premiere", command=premiere_button_clicked)
premiere_paste_button.grid(column=0, row=0)

premiere_instructions = ttk.Label(premiere_paste_frame, padding=5, text="Copy a timeline clip from Premiere\n"
                                                                        "with a keyframed mask on it")
premiere_instructions.configure(justify="center")
premiere_instructions.grid(column=0, row=1)

premiere_status = tkinter.StringVar()
premiere_status.set("Video Clip: (none)")
premiere_status_label = ttk.Label(premiere_paste_frame, padding=5, textvariable=premiere_status, wraplength=100)
premiere_status_label.config(state="disabled")
premiere_status_label.grid(column=0, row=2)

ae_paste_frame = ttk.Frame(paste_frame, padding=5, borderwidth=1, relief="solid")
ae_paste_frame.grid(column=2, row=0, sticky=(N,S))
ae_paste_button = ttk.Button(ae_paste_frame, text="Paste from After Effects", command=ae_button_clicked)
ae_paste_button.grid(column=0, row=0)

ae_instructions = ttk.Label(ae_paste_frame, padding=5, text="Run Mask_To_JSON.jsx with a keyframed\n"
                                                            "mask selected and copy the results")
ae_instructions.configure(justify="center", state="disabled")
ae_instructions.grid(column=0, row=1)

ae_status = tkinter.StringVar()
ae_status.set("Keyframes: (none)")
ae_status_label = ttk.Label(ae_paste_frame, padding=5, textvariable=ae_status, wraplength=100)
ae_status_label.config(state="disabled")
ae_status_label.grid(column=0, row=2)

process_frame = ttk.Frame(main_frame, padding=10, borderwidth=1, relief="solid")
process_frame.grid(column=0, row=1, sticky=(N,W,E,S))

process_button = ttk.Button(main_frame, text="Not Ready", command=convert_button_clicked)
process_button.grid(column=0, row=3, sticky=(N,W,E,S))
process_button.config(state=tkinter.DISABLED)


window.mainloop()