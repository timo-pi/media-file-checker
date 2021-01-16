import openpyxl
import os
import subprocess
import zipfile

# Variables
#path = '/Users/timopiechotta/Downloads/scorm/'
#unzip_path = '/Users/timopiechotta/Downloads/scorm/unzip/'
#unzipped_directories = []
#exif = 'exiftool'

# Variables
path = 'c:\\temp\\scorm\\'
unzip_path = 'c:\\temp\\scorm\\unzip\\'
unzipped_directories = []
exif = 'c:\\temp\\exiftool.exe'

# Methods
def check_file(file_input):
    print("Analyzing media file...")
    metadata = []
    process = subprocess.Popen([exif, file_input], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, universal_newlines=True, encoding='utf8', errors='ignore')
    for output in process.stdout:
        # print(output.strip())
        info = []
        line = output.strip().split(':')
        info.append(line[0].strip())
        info.append(line[1].strip())
        # metadata[line[0].strip()] = line[1].strip()
        metadata.append(info)
    #print(metadata)
    return metadata

def write_to_excel(media_data, unzip_dir):
    print("Writing media report " + unzip_dir + " to Excel.")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Info'
    for rows in media_data:
        ws.append(rows)
    wb.save('media_report.xlsx')

def filter_report(report, file_paths):
    print("Filtering media data for report.")
    data_final = [['file path', 'file size', 'file type', 'MIME Type', 'image width', 'image height', 'image size', 'Megapixels', 'media duration', 'compressor name', 'frame rate', 'avg. bitrate', 'Encoder', 'Video Frame Rate', 'Major Brand', 'Duration', 'Compressor ID', 'Track Duration', 'Compatible Brands']]
    # convert to dictionary
    file_counter = 0
    for media_file in report:
        row = []
        print('-------------------------------')
        data_dict = {}
        for data in media_file:
            data_dict[data[0]] = data[1]
            #data_dict = {data[0]: data[1]}
            print(data[0], " *** ", data[1])
            #row.append(data[0])
            #row.append(data[1])
        try:
            row.append(file_paths[file_counter])
            file_counter += 1
        except:
            print("Warning: No more file paths to grab!")
        row.append(data_dict.get('File Size')) if 'File Size' in data_dict else row.append("")
        row.append(data_dict.get('File Type')) if 'File Type' in data_dict else row.append("")
        row.append(data_dict.get('MIME Type')) if 'MIME Type' in data_dict else row.append(" ")
        row.append(data_dict.get('Image Width')) if 'Image Width' in data_dict else row.append(" ")
        row.append(data_dict.get('Image Height')) if 'Image Height' in data_dict else row.append(" ")
        row.append(data_dict.get('Image Size')) if 'Image Size' in data_dict else row.append(" ")
        row.append(data_dict.get('Megapixels')) if 'Megapixels' in data_dict else row.append(" ")
        row.append(data_dict.get('Media Duration')) if 'Media Duration' in data_dict else row.append(" ")
        row.append(data_dict.get('Compressor Name')) if 'Compressor Name' in data_dict else row.append(" ")
        row.append(data_dict.get('Video Frame Rate')) if 'Video Frame Rate' in data_dict else row.append(" ")
        row.append(data_dict.get('Avg Bitrate')) if 'Avg Bitrate' in data_dict else row.append(" ")
        row.append(data_dict.get('Encoder')) if 'Encoder' in data_dict else row.append(" ")
        row.append(data_dict.get('Video Frame Rate')) if 'Video Frame Rate' in data_dict else row.append(" ")
        row.append(data_dict.get('Major Brand')) if 'Major Brand' in data_dict else row.append(" ")
        row.append(data_dict.get('Duration')) if 'Duration' in data_dict else row.append(" ")
        row.append(data_dict.get('Compressor ID')) if 'Compressor ID' in data_dict else row.append(" ")
        row.append(data_dict.get('Track Duration')) if 'Track Duration' in data_dict else row.append(" ")
        row.append(data_dict.get('Compatible Brands')) if 'Compatible Brands' in data_dict else row.append(" ")
        #row.append(data_dict.get('')) if '' in data_dict else row.append(" ")

        data_final.append(row)
        print(row)
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
    file_paths = []
    for folder, subfolders, filenames in os.walk(unzip_dir):
        for file in filenames:
            # print(file)
            if file.endswith('.png') or file.endswith('.jpg') or file.endswith('.gif') or file.endswith('.jpeg') or file.endswith('mp3'):
                print(os.path.join(folder, file))
                report.append(check_file(os.path.join(folder, file)))
                file_paths.append(os.path.join(folder, file))
            elif file.endswith('.mpeg') or file.endswith('.mp4') or file.endswith('.mov') or file.endswith('.avi'):
                print(os.path.join(folder, file))
                file_paths.append(os.path.join(folder, file))
                report.append(check_file(os.path.join(folder, file)))

    # create report for each unzipped SCORM package
    report_filtered = filter_report(report, file_paths)
    os.chdir(unzip_dir)
    write_to_excel(report_filtered, unzip_dir)
    print(report)