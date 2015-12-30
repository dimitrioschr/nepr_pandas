def fun1():

    import read_new_nepr_file
    import os
    import pandas
    pandas.set_option('display.width', 200)

    def get_files_in_folder(path):
        listing = os.listdir(path)
        return [os.path.join(path, l) for l in listing if os.path.isfile(os.path.join(path, l))]

    def get_folders_in_folder(path):
        listing = os.listdir(path)
        return [os.path.join(path, l) for l in listing if not os.path.isfile(os.path.join(path, l))]

    path = 'C:\\Users\\tech5\\Google Drive\\NEPR Actual'

    file_list = []
    for folder in get_folders_in_folder(path):
        for file in get_files_in_folder(folder):
            if '(' in file and '.xlsx' in file:
                file_list.append(file)


    df = read_new_nepr_file.read_nepr_file(file_list[0])
    for f in file_list[1:]:
        print(f)
        df = df.append(read_new_nepr_file.read_nepr_file(f), ignore_index=True)

    return df

if __name__ == '__main__':
    fun1()