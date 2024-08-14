class name_validations:
    def __checkValid__(self, name):
        self.name = str(name)
        self.valid_name = False
        self.valid_chars = [
            "llc",
            "ltd",
            "co.",
            "company",
            "co",
            "company.",
            "inc.",
            "inc",
        ]
        for index in self.valid_chars:
            if index in self.name.lower():
                self.valid_name = True
                return self.valid_name
        else:
            return False


if __name__ == "__main__":
    validation_obj = name_validations()
    validation_obj.__checkValid__()
