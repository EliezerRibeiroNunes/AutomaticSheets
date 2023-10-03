import json

class DBController:

    def update_database_value(self, new_value, key_data):
        
        try:
            with open('database.json', "r") as database:
                data = json.load(database)
            
            data[key_data] = new_value
            
            with open('database.json', "w") as database:
                json.dump(data, database, indent=4)
           
        except Exception as e:
            print(f"ERROR: {str(e)}")
            
    def get_all_data(self):
        try:
            with open('database.json', "r") as database:
                return json.load(database)
    
        except Exception as e:
            print(f"ERROR: {str(e)}")