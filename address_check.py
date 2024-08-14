import json


class address_validations:
    def __init__(self, address, city, state):
        self.data = {}
        self.temp = {}
        self.temp[address] = {"city": city, "state": state}
        # Load JSON data from file
        try:
            with open("address_data.json", "r") as infile:  #  if the file is available
                self.data = json.load(infile)
        except json.JSONDecodeError:                       # if the file is not available
            with open("address_data.json", "w") as outfile:
                json.dump({}, outfile)
        except FileNotFoundError:
            with open("address_data.json", "w") as outfile:
                json.dump({}, outfile)
        for key in self.temp.keys():
            if key in self.data.keys():
                # print(f"Data is already available {self.data}", end="\n")
                # print("Phone Numbers are added")
                #  Clear the data dictionary
                self.data.clear()
                self.data.update(self.temp)
                self._save_to_json()
                return True
            else:
                self.data.clear()
                self.data.update(self.temp)
                # Save the data to a JSON file
                self._save_to_json()
                return False

    def _save_to_json(self):  # / Save method
        with open("address_data.json", "w") as outfile:
            json.dump(self.data, outfile)


if __name__ == "__main__":

    checking_obj = address_validations()
    # Update the data
    checking_obj._save_to_json()