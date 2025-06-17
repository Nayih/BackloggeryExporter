import re, os, requests, json, sys, xlsxwriter, datetime

def sort_game_data(games_dict):

    def sort_games(games_list):
        result = []
        parent_dict = {}
        standalone_games = []

        for game in games_list:
            if game["Parent ID"]:
                parent_dict.setdefault(game["Parent ID"], []).append(game)
            else:
                standalone_games.append(game)

        standalone_games.sort(key=lambda x: x["Title"])

        for game in standalone_games:
            result.append(game)

            if game["Is Parent"]:
                children = parent_dict.get(game["ID"], [])
                children.sort(key=lambda x: x["Title"])
                result.extend(children)

        return result

    organized_dict = {}
    for platform, games_list in games_dict.items():
        organized_dict[platform] = sort_games(games_list)

    return organized_dict

def n_cls():
    os.system('cls' if os.name=='nt' else 'clear')

def fetch(URL, payload):
    try:
        response = requests.post(URL, data=json.dumps(payload))
    except Exception as e:
        print(f"[Error]: Backup not performed.\n\n[Log]: {e}\n")
        sys.exit()

    return str(response.text)

def validate_profile(username):
    response = fetch("https://backloggery.com/api/fetch_profile.php", {"type": "load_user_profile", "username": username})

    if (str(response) == "False"):
        print("[Error]: The username provided doesn't match a valid account, or the profile is private.")
        sys.exit()

def extract_data(content, patterns, unique=False, startTag=None, endTag=None, flags=None):
    
    #Nayoh Pattern Extractor Tutorial:

    #Content: The text to be analyzed.
    #Paterns: The pattern to be found, Example: [r'<option value="(.*?)">(.*?)</option>']
    #Unique: Will ensure that what you are looking for is found or the program will terminate. It will also find and return only the first occurrence it finds, useful if used with startTag and endTag in complex texts or if you only want 1 single return.
    #StartTag: Will cut the content at the beginning when finding a pattern. Example: startTag= r'<h1>' (Valid only for 1 occurrence)
    #EndTag: Will cut the content at the end when finding a pattern. Example: endTag= r'</h1>' (Valid only for 1 occurrence)
    #flags: Regular expression frags: re.A, re.I, re.L, re.M, re.S, re.U, re.X
    
    if startTag:
        startPosition = content.find(startTag)

        if startPosition == -1:
            print("[Error]: Unable to get data from extract_data. (startPosition -1)\n")
            sys.exit()

        content = content[startPosition:]
    
    if endTag:
        endPosition = content.find(endTag)

        if endPosition == -1:
            print("[Error]: Unable to get data from extract_data. (endPosition -1)\n")
            sys.exit() 

        if endPosition: content = content[:endPosition + len(endTag)]
    
    if flags:
        compiled_patterns = [re.compile(pattern, flags) for pattern in patterns]
    else:
        compiled_patterns = [re.compile(pattern) for pattern in patterns]

    matches = []
    
    for pattern in compiled_patterns:
        if unique:
            match = re.search(pattern, content)
            if match:
                match = match.group(1)
                matches.append(match)
            else:
                print("[Error]: Unable to get game data from extract_data. (match unique = None)\n")
                sys.exit()
        else:
            match = re.findall(pattern, content)
            matches.extend(match)

    return matches

def unescape_json(text):
    return json.loads(f'"{text}"')

def convert_value(value, type):
    try:
        return type(value)
    except ValueError:
        return value

def get_game_data(library):
    switch_mappings = {
        "region": {
            "0": "", "1": "Free", "2": "North America", "3": "Japan", "4": "PAL", "5": "China", "6": "Korea", "7": "Brazil", "8": "Asia"
            },
        "format": {
            "0": "", "1": "Digital", "20": "Physical", "21": "Physical (Game Only)", "22": "Physical (Incomplete)", "28": "Physical (Complete In Box)", "29": "Physical (Sealed)", "30": "Physical (Licensed Repro)", "31": "Physical (Unlicensed Repro)"
            },
        "owner": {
            "0": "", "1": "Own", "7": "Wishlist", "5": "Household", "6": "Subscription", "3": "Played It", "2": "Formerly Owned", "4": "Other"
            },
        "status": {
            "10": "Unplayed", "20": "Unfinished", "30": "Beaten", "40": "Completed", "60": "Endless", "80": ""
            },
        "priority": {
            "80": "Now Playing", "70": "Ongoing", "60": "Paused", "50": "High", "40": "Normal", "30": "Low", "20": "Replay", "10": "Shelved"
            },
        "rating": {
            "None": "0", "1": "0.5", "2": "1", "3": "1.5", "4": "2", "5": "2.5", "6": "3", "7": "3.5", "8": "4", "9": "4.5", "10": "5"
            },
        "difficulty": {
            "0": "", "20": "Relaxing", "10": "Too Easy", "50": "Too Hard", "60": "Unfair", "30": "Moderate", "40": "Hurts So Good"
        }
    }

    game_data = {}
    len_library = len(library)
    count = 0

    for item in library:
        n_cls()
        count += 1
        print(f"Getting Data... ({count}/{len_library})")

        game_info = fetch("https://backloggery.com/api/fetch_gameinfo.php", {"game_inst_id": item[0]})
        extra_info = extract_data(game_info, [r'"title":"(.*?)",".*?"notes":"(.*?)","achieve_score":(.*?),"achieve_total":(.*?),"online_info":"(.*?)",".*?"review":"(.*?)","rating":(.*?),"difficulty":(.*?)}'])
        
        region = switch_mappings["region"].get(item[3], "Unknown")
        game_format = switch_mappings["format"].get(item[4], "Unknown")
        owner = switch_mappings["owner"].get(item[5], "Unknown")
        status = switch_mappings["status"].get(item[1], "Unknown")
        priority = switch_mappings["priority"].get(item[2], "Unknown")
        rating = switch_mappings["rating"].get(extra_info[0][6], "Unknown")
        difficulty = switch_mappings["difficulty"].get(extra_info[0][7], "Unknown")
        
        title = unescape_json(extra_info[0][0])
        notes = unescape_json(extra_info[0][1])
        online_info = unescape_json(extra_info[0][4])
        review = unescape_json(extra_info[0][5])
        
        sub_platform = item[7][1:-1] if item[7] != "null" else ""
        parent_id = item[8] if item[8] != "null" else ""
        achievement_score = extra_info[0][2] if extra_info[0][2] != "null" else ""
        achievement_total = extra_info[0][3] if extra_info[0][3] != "null" else ""
        is_parent = item[9] == "1"
        
        id = convert_value(item[0], int)
        parent_id = convert_value(parent_id, int)
        achievement_score = convert_value(achievement_score, int)
        achievement_total = convert_value(achievement_total, int)
        rating = convert_value(rating, float)
        
        if achievement_score == 0 and achievement_total == 0:
            achievement_score = ""
            achievement_total = ""
        
        temp_dict = {
            "ID": id, "Title": title, "Platform": item[6], "Sub Platform": sub_platform, "Region": region,
            "Format": game_format, "Ownership": owner, "Achievements Score": achievement_score,
            "Achievements Total": achievement_total, "Status": status, "Priority": priority,
            "Parent ID": parent_id, "Is Parent": is_parent, "Online Information": online_info,
            "Star Rating": rating, "Difficulty Rating": difficulty, "Notes": notes, "Review": review
        } 
        
        category = temp_dict["Platform"]
        if category not in game_data:
            game_data[category] = []
            
        game_data[category].append(temp_dict)

    return game_data

def saveXlsx(dict_data):
    def create_formats(workbook, bg_color, font_color, border=1, border_color="#000000", bold=True):
        base_format = {
            'border': border,
            'border_color': border_color,
            'bg_color': bg_color,
            'font_color': font_color,
            'bold': bold,
            'text_wrap': True,
        }
        return {
            'normal': workbook.add_format(base_format),
            'aligned': workbook.add_format({**base_format, 'align': 'center'})
        }

    column_widths = [
        ('A', 13), ('B', 50), ('C', 16), ('D', 16), ('E', 18), ('F', 15), ('G', 20), ('H', 21), ('I', 16),
        ('J', 15), ('K', 13), ('L', 15), ('M', 24), ('N', 2.57), ('O', 2.57), ('P', 2.57), ('Q', 2.57),
        ('R', 2.57), ('S', 19), ('T', 50), ('U', 50)
    ]

    date = datetime.datetime.now().strftime("%Y-%m-%d")
    workbook = xlsxwriter.Workbook(f'[Backup]Backloggery ({date}).xlsx')

    col_format = workbook.add_format({
        'bold': True, 'align': 'center', 'bg_color': '#525252', 'font_color': '#ffffff',
        'border': 6, 'border_color': '#000000', 'italic': True
    })

    row_formats = {
        "default": create_formats(workbook, bg_color="#493A3A", font_color="#ffffff", bold=False),
        "father": create_formats(workbook, bg_color="#6D5858", font_color="#ffffff"),
        "child": create_formats(workbook, bg_color="#BDACAC", font_color="#493A3A", bold=False),
        "collection": create_formats(workbook, bg_color="#685877", font_color="#ffffff"),
        "collection_child": create_formats(workbook, bg_color="#A596B1", font_color="#332C3A", bold=False),
    }

    green = workbook.add_format({'font_color': '#9BBB59'})
    red = workbook.add_format({'font_color': '#C0504D'})

    for category, data in dict_data.items():
        worksheet = workbook.add_worksheet(category[:31])

        worksheet.conditional_format('$L2:$L1048576', {
            'type': 'cell', 'criteria': 'equal to', 'value': 'TRUE', 'format': green
        })
        worksheet.conditional_format('$L2:$L1048576', {
            'type': 'cell', 'criteria': 'equal to', 'value': 'FALSE', 'format': red
        })
        worksheet.conditional_format('$N$2:$R$1048576', {
            'type': 'icon_set', 'icon_style': '3_traffic_lights_rimmed',
            'icons': [
                {'criteria': '>=', 'type': 'percent', 'value': 67},
                {'criteria': '>=', 'type': 'percent', 'value': 33}
            ]
        })

        headers = [k for k in data[0].keys() if k != 'Achievements Total']
        control = 0

        for col_num, (index, width) in enumerate(column_widths):
            worksheet.set_column(f'{index}:{index}', width)
            if col_num >= 18:
                worksheet.write(0, col_num, headers[col_num - 4], col_format)
            elif 14 <= col_num <= 17:
                control += 1
                if control == 4:
                    worksheet.merge_range(0, 13, 0, 17, "Star Rating", col_format)
            else:
                worksheet.write(0, col_num, headers[col_num], col_format)

        for row_num, register in enumerate(data, start=1):
            worksheet.set_row(row_num, 15.1)
            if register["Is Parent"] and register["Status"] != "":
                format_type = "father"
            elif register["Is Parent"] and register["Status"] == "":
                format_type = "collection"
            elif register["Parent ID"] != "":
                parent_id = register["Parent ID"]
                father = next((f for f in data if f["ID"] == parent_id), {})
                format_type = "collection_child" if father.get("Status", "") == "" else "child"
            else:
                format_type = "default"

            for col_num, header in enumerate(headers):
                if col_num == 7:
                    achievements_total = dict_data[category][row_num - 1].get('Achievements Total', 0)
                    worksheet.conditional_format(row_num, col_num, row_num, col_num, {
                        'type': 'data_bar', 'bar_color': '#638EC6', 'bar_solid': True,
                        'bar_border_color': '#000000', 'min_type': 'num', 'max_type': 'num',
                        'min_value': 0, 'max_value': achievements_total
                    })

                    bar_fmt = {
                        "default": {'bg_color': '#493A3A', 'font_color': '#ffffff'},
                        "father": {'bg_color': '#6D5858', 'font_color': '#ffffff', 'bold': True},
                        "child": {'bg_color': '#BDACAC', 'font_color': '#493A3A'},
                        "collection": {'bg_color': '#685877', 'font_color': '#ffffff', 'bold': True},
                        "collection_child": {'bg_color': '#A596B1', 'font_color': '#332C3A'}
                    }[format_type]
                    
                    data_bar_format = workbook.add_format({**bar_fmt, 'text_wrap': True, 'align': 'center', 'border': 1, 'border_color': '#000000'})
                    data_bar_format.set_num_format(f'#,##0" / {achievements_total}"')
                    worksheet.write(row_num, col_num, register[header], data_bar_format)

                elif col_num == 13:
                    rating = register[header]
                    for i in range(5):
                        formula = ("" if rating == "Unknown" else
                                   f'=IF({rating}>={i+1},1,IF(INT({rating})={i},MOD({rating},1),0))')
                        worksheet.write(row_num, col_num + i, "" if rating == "Unknown" else formula, row_formats[format_type]['aligned'])

                elif col_num == 0 or col_num in range(2, 7) or col_num in range(8, 12) or col_num == 14:
                    target_col = col_num + 4 if col_num >= 14 else col_num
                    worksheet.write(row_num, target_col, register[header], row_formats[format_type]['aligned'])
                else:
                    target_col = col_num + 4 if col_num >= 14 else col_num
                    worksheet.write(row_num, target_col, register[header], row_formats[format_type]['normal'])

    workbook.close()





























