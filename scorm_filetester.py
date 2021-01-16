import openpyxl
import os
import subprocess
import zipfile

# Variables
path = 'c:\\temp\\scorm\\'
unzip_path = 'c:\\temp\\scorm\\unzip\\'
unzipped_directories = []
exif = 'c:\\temp\\exiftool.exe'

# Methods
def check_file(file_input):
    print("Analyzing media file...")
    metadata = []
    try:
        process = subprocess.Popen([exif, file_input], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, universal_newlines=True)
        for output in process.stdout:
            # print(output.strip())
            info = []
            line = output.strip().split(':')
            info.append(line[0].strip())
            info.append(line[1].strip())
            # metadata[line[0].strip()] = line[1].strip()
            metadata.append(info)
    except:
        print('Error checking file with exif-tool: ' + file_input)
    print(metadata)
    return metadata

def write_to_excel(media_data, unzip_dir):
    print("Writing media report " + unzip_dir + " to Excel.")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Info'
    for rows in media_data:
        ws.append(rows)
    wb.save('media_report.xlsx')

def filter_report(report):
    print("Filtering media data for report.")
    data_final = [['file name', ' file size', 'directory', 'file type', 'image width', 'image height', 'image size', 'media duration', 'compressor name', 'frame rate', 'avg. bitrate']]
    # convert to dictionary

    for media_file in report:
        row = []
        print('-------------------------------')
        for data in media_file:
            data_dict = {data[0]: data[1]}
            print(data[0], data[1])
            if 'File Name' in data_dict: row.append(data_dict.get('File Name'))
            if 'File Size' in data_dict: row.append(data_dict.get('File Size'))
            if 'Directory' in data_dict: row.append(data_dict.get('Directory'))
            if 'File Type' in data_dict: row.append(data_dict.get('File Type'))
            if 'Image Width' in data_dict: row.append(data_dict.get('Image Width'))
            if 'Image Height' in data_dict: row.append(data_dict.get('Image Height'))
            if 'Image Size' in data_dict: row.append(data_dict.get('Image Size'))
            if 'Media Duration' in data_dict: row.append(data_dict.get('Media Duration'))
            if 'Compressor Name' in data_dict: row.append(data_dict.get('Compressor Name'))
            if 'Video Frame Rate' in data_dict: row.append(data_dict.get('Video Frame Rate'))
            if 'Avg Bitrate' in data_dict: row.append(data_dict.get('Avg Bitrate'))
        data_final.append(row)
        print(data_final)
    return data_final

# **** Script ****
# Unzip all zip-files in unzip directory
for file in os.listdir(path):
    if file.endswith('.zip'):
        # print("Unzipping " + file)
        with zipfile.ZipFile(path + file, 'r') as zip_ref:
            new_directory = unzip_path + file[:-4]
            zip_ref.extractall(new_directory)
            unzipped_directories.append(new_directory)

# find all images and videos of unzipped content
for unzip_dir in unzipped_directories:
    # print("Searching for all media files in " + unzip_dir)
    report = []
    for folder, subfolders, filenames in os.walk(unzip_dir):
        for file in filenames:
            # print(file)
            if file.endswith('.png') or file.endswith('.jpg') or file.endswith('.mp4'):
                # print(os.path.join(folder, file))
                report.append(check_file(os.path.join(folder, file)))
                report.append(os.path.join(folder, file))
            elif file.endswith('.mpeg') or file.endswith('.mp4'):
                # print(os.path.join(folder, file))
                report.append(check_file(os.path.join(folder, file)))
                report.append(os.path.join(folder, file))
    print(report)

    # create report for each unzipped SCORM package
    report_filtered = filter_report(report)
    os.chdir(unzip_dir)
    write_to_excel(report_filtered, unzip_dir)
    # print(report)
