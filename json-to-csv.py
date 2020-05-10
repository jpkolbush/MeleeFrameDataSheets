import os, json, csv
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

JSON_PATH = 'framedata-json/'
MOVE_DICT = {
    'jab1': "Jab 1",
    'jab2': "Jab 2",
    'jab3': "Jab 3",
    'rapidjabs_start': "Rapid Jabs Startup",
    'rapidjabs_loop': "Rapid Jabs Loop",
    'rapidjabs_end': "Rapid Jabs End",
    'dashattack': "Dash Attack",
    'ftilt_m': "F-tilt",
    'utilt': "U-tilt",
    'dtilt': "D-tilt",
    'fsmash_m': "F-Smash",
    'usmash': "U-Smash",
    'dsmash': "D-Smash",
    'nair': "Nair",
    'fair': "Fair",
    'bair': "Bair",
    'uair': "Uair",
    'dair': "Dair",
    'grab': "Grab",
    'dashgrab': "Dash Grab",
    'pummel': "Pummel",
    'fthrow': "F-Throw",
    'bthrow': "B-Throw",
    'uthrow': "U-Throw",
    'dthrow': "D-Throw"
}

def generate_row(move_name, move_json_data):
    row = []
    row.append(MOVE_DICT[move_name])
    hits = []
    hit_frames_list = move_json_data.get("hitFrames")
    for hit_frame in hit_frames_list:
        hits.append("{}-{}".format(hit_frame["start"], hit_frame["end"]))
    row.append(", ".join(hits))
    row.append(move_json_data.get("totalFrames"))
    row.append(move_json_data.get("iasa"))
    row.append(move_json_data.get("landingLag"))
    row.append(move_json_data.get("lcancelledLandingLag"))
    notes = ""
    if (move_json_data.get("autoCancelBefore") and move_json_data.get("autoCancelAfter")):
        notes +=  "Autocancel <{} {}>".format(move_json_data.get("autoCancelBefore"), move_json_data.get("autoCancelAfter"))
    elif (move_json_data.get("autoCancelBefore")):
        notes += "Autocancel <{} ".format(move_json_data.get("autoCancelBefore"))
    elif (move_json_data.get("autoCancelAfter")):
        notes += "Autocancel {}>".format(move_json_data.get("autoCancelAfter"))
    
    if (move_json_data.get("chargeFrame")):
        notes += "Charge Frame: {}".format(move_json_data.get("chargeFrame"))

    # res["Notes"] = notes
    row.append(notes)

    damage = []
    throw = move_json_data.get("throw")
    if throw:
        print (move_name)
        print (throw)
        damage.append(str(throw["damage"]))
    else:
        hitboxes_list = move_json_data.get("hitboxes")
        for hit_box in hitboxes_list:
            damage.append(str(hit_box["damage"]))
    # res["Base Damage"] = ", ".join(damage)
    row.append(", ".join(damage))

    return row

def generate_sheet(character_json, workbook):
    character_name = character_json.split(".")[0]
    print(character_name)
    with open(JSON_PATH + character_json) as f:
        json_data = json.load(f)
    
    rows = []
    for key in json_data.keys():
        if (key != "ftilt_h" and key != "ftilt_mh" and key != "ftilt_ml" and key != "ftilt_l" 
            and  key != "fsmash_h" and key != "fsmash_mh"  and key != "fsmash_ml"  and key != "fsmash_l"
            and json_data[key] != None and key[0] != '0' and 'special' not in key):
            rows.append(generate_row(key, json_data[key]))

    sheet = workbook.create_sheet(character_name)
    header = ['Move', "Active Hits", "Total Frames", "IASA", "Landing Lag", "L-Cancelled", "Notes", "Base Damage"]
    sheet.append(header)
    for row in rows:
        sheet.append(row)

    return workbook
def main():
    wb = Workbook()
    json_files = os.listdir(JSON_PATH)
    print(json_files)
    json_files.sort()
    for filename in json_files:
        wb = generate_sheet(filename, wb)
    wb.save("FrameData.xlsx")
    


if __name__ == '__main__':
    main()