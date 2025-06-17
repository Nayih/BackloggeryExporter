from functions import *

def start():
    game_data = {}
    n_cls()

    print("Enter the username of the account from which the data will be extracted. (Profile must be public!)")
    username = input("Username: ")
    
    validate_profile(username)
    
    fetch_library = fetch("https://www.backloggery.com/api/fetch_library.php", {"type": "load_user_library","username": username})
    library = extract_data(fetch_library, [r'"game_inst_id":(\d+),".*?"status":(\d+),"priority":(\d+),"region":(\d),".*?"phys_digi":(\d+),"own":(\d),".*?"platform_title":"(.*?)","sub_platform_title":(.*?),".*?"parent_inst_id":(.*?),"is_parent":(\d)}'])
    
    game_data = get_game_data(library)
    sorted_data = sort_game_data(game_data)
    
    saveXlsx(sorted_data)

start()