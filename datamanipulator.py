from xlsxfile import XLSXFile


class DataManipulator():

    def __init__(self) -> None:
        pass



    def extract_data(self, file: XLSXFile, start: float, stop: float) -> dict:
        time = file.data["time"]
        extracted_time = self.extract_time_data(data=time, start=start, stop=stop)
        indexes = self.find_start_and_stop_indexes(data=time, start_value=extracted_time[0], stop_value=extracted_time[len(extracted_time) - 1])

        output = {
            "start_time": start,
            "stop_time": stop,
            "time": extracted_time,
        }

        for d in file.data["channels"]:
            key = d["type"]
            detector_data = file.data["data_points"][key]
            extracted_data = self.extract_from_data_on_indexes(data=detector_data, start_index=indexes[0], stop_index=indexes[1])
            output[key] = extracted_data
        
        return output





    ### Data extraction method.
    def extract_time_data(self, data: list, start: float, stop: float) -> list:
        return [point for point in data if point >= start if point <= stop]


    ### Method returs extracted data from a list based on specified indexes.
    def extract_from_data_on_indexes(self, data: list, start_index: int, stop_index: int) -> list:
        if stop_index < start_index or len(data) - 1 < stop_index or start_index < 0:
            print(f"Fundamentally wrong indexes were passed to extract data, are you sure {start_index} and {stop_index} are correct?")
            exit()

        return data[start_index : stop_index + 1]
    

    ### Finds start and stop indexes of the extracted data using extract_time_data method.
    def find_start_and_stop_indexes(self, data: list, start_value: float, stop_value: float) -> tuple:
        start_index = -1
        stop_index = -1
        index = 0

        for point in data:
            if point == start_value:
                start_index = index
            if point == stop_value:
                stop_index = index
                break
            
            index += 1
        
        return (start_index, stop_index)


    
