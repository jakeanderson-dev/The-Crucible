from pymongo import MongoClient
import cv2
import argparse
import subprocess
import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows
import requests

# Parse command-line arguments
def parse_arguments():
    parser = argparse.ArgumentParser(description="All sorts of stuff.")
    parser.add_argument("--baselight", type=str, help="Path to baselight export file.")
    parser.add_argument("--xytech", type=str, help="Path to xytech file.")
    parser.add_argument("--process", type=str, help="Path to video file.")
    parser.add_argument("--output_xls", type=str, help="Output XLS.")
    parser.add_argument("--thumbnails", type=str, help="Directory to save thumbnails.", default="thumbnails")
    args = parser.parse_args()
    return args

# Read file and return its contents
def read_file(file_path):
    try:
        with open(file_path, 'r') as file:
            contents = file.readlines()
        return contents
    except FileNotFoundError:
        print(f"The file '{file_path}' does not exist.")
        exit(1)
    except Exception as e:
        print(f"An error occurred while reading the file: {e}")
        exit(1)

# Process baselight data from file
def process_baselight_data(lines):
    data = []
    for line in lines:
        if line.strip():
            parts = line.strip().split()
            folder = parts[0]
            frames = [int(frame) for frame in parts[1:] if frame.isdigit()]
            data.append({"Folder": folder, "Frames": frames})
    return data

# Process xytech data from file
def process_xytech_data(lines):
    xytech_data = {}
    header_info = {}
    notes = ""
    capturing_notes = False
    header_info['Workorder'] = lines[0].strip().split()[-1]
    header_info['Producer'] = lines[2].split(': ')[1].strip()
    header_info['Operator'] = lines[3].split(': ')[1].strip()
    header_info['Job'] = lines[4].split(': ')[1].strip()
    
    for line in lines[6:]:
        if 'Notes:' in line:
            capturing_notes = True
            continue
        if capturing_notes:
            notes += line.strip() + " "
        else:
            path = line.strip()
            if path:
                xytech_data[path] = []
                
    return xytech_data, header_info, notes.strip()

# Connect to MongoDB
def connect_to_mongo():
    client = MongoClient('mongodb+srv://quackis:Password1234@poo.73zpvh2.mongodb.net/')
    db = client['poo']
    return db

# Process video to extract duration, total frames, and FPS
def process_video(file_path):
    cap = cv2.VideoCapture(file_path)
    if not cap.isOpened():
        print("Could not open video file.")
        return None, None

    duration = cap.get(cv2.CAP_PROP_FRAME_COUNT) / cap.get(cv2.CAP_PROP_FPS)
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    fps = cap.get(cv2.CAP_PROP_FPS)
    cap.release()
    return duration, total_frames, fps

# Find records in MongoDB that match the frame criteria
def find_records_in_range(db, total_frames, fps, xytech_data, video_path, thumbnail_dir):
    records = {}
    query = {"Frames": {"$elemMatch": {"$lte": total_frames}}}
    results = db.baselight.find(query)
    
    for record in results:
        mapped_path = map_path_to_xytech(xytech_data, record['Folder'])
        for frame in record['Frames']:
            if frame <= total_frames:
                if mapped_path not in records:
                    records[mapped_path] = []
                records[mapped_path].append(frame)
                
    return process_frame_ranges(records, fps, video_path, thumbnail_dir)

# Convert frame numbers to timecode
def frames_to_timecode(frames, fps):
    hours = int(frames / (fps * 3600))
    minutes = int((frames % (fps * 3600)) / (fps * 60))
    seconds = int((frames % (fps * 60)) / fps)
    frame_count = int(frames % fps)
    return f"{hours:02}:{minutes:02}:{seconds:02}:{frame_count:02}"

# Export data to an Excel file with embedded images
def export_to_xls(data, header_info, notes, output_path, thumbnail_dir):
   
    header_info['Notes'] = notes
    header_df = pd.DataFrame([header_info], columns=['Workorder', 'Producer', 'Operator', 'Job', 'Notes'])
    data_df = pd.DataFrame(data, columns=['Location', 'Frame Range', 'Timecode Range', 'Thumbnail'])

    wb = Workbook()
    ws = wb.active

    for r in dataframe_to_rows(header_df, index=False, header=True):
        ws.append(r)

    ws.append([])
    ws.append(['Location', 'Frame Range', 'Timecode Range', 'Thumbnail']) 
    
    row_index = ws.max_row + 1
    for index, row in data_df.iterrows():
        ws.append(list(row.values)[:3])  # Excluding Thumbnail path
        thumbnail_path = os.path.join(thumbnail_dir, os.path.basename(row['Thumbnail']))
        if os.path.exists(thumbnail_path):
            img = Image(thumbnail_path)
            img.width, img.height = 96, 74  # Adjust as necessary for your dimensions
            cell = f'D{row_index}'  # Column D for thumbnails
            ws.add_image(img, cell)
        row_index += 1

    wb.save(output_path)
    print("Data processed and exported to XLS with images embedded.")

# Map baselight path to xytech path
def map_path_to_xytech(xytech_data, baselight_path):
    last_four_segments = '/'.join(baselight_path.split('/')[-4:])
    for xytech_path in xytech_data:
        if last_four_segments in xytech_path:
            return xytech_path
    return baselight_path

# Process frame ranges to prepare for rendering
def process_frame_ranges(records, fps, video_path, thumbnail_dir):
    grouped_data = []
    for location, frames in records.items():
        frames = sorted(set(frames)) 
        if len(frames) < 2:
            continue 

        range_start = frames[0]
        last_frame = frames[0]
        start_timecode = frames_to_timecode(range_start, fps)

        for i, frame in enumerate(frames[1:]):
            if frame == last_frame + 1:
                last_frame = frame
            else:
                if last_frame != range_start:
                    end_timecode = frames_to_timecode(last_frame, fps)
                    middle_frame = (range_start + last_frame) // 2
                    thumbnail_filename = f"thumbnail_{middle_frame}.png"
                    thumbnail_path = capture_thumbnail(video_path, middle_frame, thumbnail_dir, thumbnail_filename)
                    grouped_data.append({
                        'Location': location,
                        'Frame Range': f"{range_start}-{last_frame}",
                        'Timecode Range': f"{start_timecode}-{end_timecode}",
                        'Thumbnail': thumbnail_path
                    })
                range_start = frame
                last_frame = frame
                start_timecode = frames_to_timecode(range_start, fps)

        if last_frame != range_start:
            end_timecode = frames_to_timecode(last_frame, fps)
            middle_frame = (range_start + last_frame) // 2
            thumbnail_filename = f"thumbnail_{middle_frame}.png"
            thumbnail_path = capture_thumbnail(video_path, middle_frame, thumbnail_dir, thumbnail_filename)
            grouped_data.append({
                'Location': location,
                'Frame Range': f"{range_start}-{last_frame}",
                'Timecode Range': f"{start_timecode}-{end_timecode}",
                'Thumbnail': thumbnail_path
            })

    return grouped_data

# Capture thumbnail for a specific frame in a video
def capture_thumbnail(video_path, frame_number, thumbnail_dir, filename):
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        print(f"Could not open video file {video_path}.")
        return None
    cap.set(cv2.CAP_PROP_POS_FRAMES, frame_number)
    success, frame = cap.read()
    if not success:
        print(f"Could not read frame at {frame_number}.")
        cap.release()
        return None
    thumbnail = cv2.resize(frame, (96, 74))
    if not os.path.exists(thumbnail_dir):
        os.makedirs(thumbnail_dir)
    thumbnail_path = os.path.join(thumbnail_dir, filename)
    cv2.imwrite(thumbnail_path, thumbnail)
    cap.release()
    return thumbnail_path

# Render video clips based on the data provided
def render_video_clips(data, video_path, output_folder, fps):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for item in data:
        start_frame, end_frame = map(int, item['Frame Range'].split('-'))
        start_timecode, end_timecode = item['Timecode Range'].split('-')
        start_time_ms = timecode_to_ms(start_timecode, fps)
        end_time_ms = timecode_to_ms(end_timecode, fps)
        
        output_filename = f"render_{start_frame}_{end_frame}.mp4"
        output_path = os.path.join(output_folder, output_filename)
        
        cmd = f"ffmpeg -ss {start_time_ms / 1000} -i {video_path} -t {(end_time_ms - start_time_ms) / 1000} -c copy {output_path}"
        subprocess.run(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)


# Convert milliseconds to timecode
def milliseconds_to_timecode(milliseconds):
    seconds = milliseconds / 1000
    hours = int(seconds // 3600)
    seconds %= 3600
    minutes = int(seconds // 60)
    seconds %= 60
    milliseconds = int((seconds - int(seconds)) * 1000)
    seconds = int(seconds)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}.{milliseconds:03d}"

# Upload rendered video clips to Frame.io
def upload_to_frameio(api_token, project_id, folder_path):
    url = f"https://api.frame.io/v2/projects/{project_id}/assets"
    headers = {
        "Authorization": f"Bearer {api_token}"
    }
    
    for filename in os.listdir(folder_path):
        if filename.endswith(".mp4"):
            file_path = os.path.join(folder_path, filename)
            data = {
                "name": filename,
                "type": "file",
                "filetype": "video/mp4",
                "parent_id": None  # Make sure this is correct or necessary
            }
            response = requests.post(url, headers=headers, json=data)
            if response.status_code == 201:
                upload_url = response.json().get('upload_url')
                file_data = open(file_path, 'rb').read()
                upload_response = requests.put(upload_url, headers={"Content-Type": "video/mp4"}, data=file_data)
                if upload_response.status_code == 200:
                    print(f"Uploaded {filename} to Frame.io")
                else:
                    print(f"Failed to upload {filename}: {upload_response.status_code} {upload_response.text}")
            else:
                print(f"Failed to create Frame.io asset: {response.status_code} {response.text}")

# Convert timecode to milliseconds
def timecode_to_ms(timecode, fps):
    hours, minutes, seconds, frames = map(int, timecode.split(':'))
    total_seconds = hours * 3600 + minutes * 60 + seconds + frames / fps
    return int(total_seconds * 1000)

# Main function
def main():
    args = parse_arguments()
    db = connect_to_mongo()
    api_token = "fio-u-M-s00cTTmhoj1cMtMPhDqhHuv9a33MnctDXjDGjkDTj7u8fHYJ_R_XUWxdmNNyRG"
    project_id = "07999592-9875-409b-b620-35df2477394b"
    renders_folder = "renders"

    if args.baselight and args.xytech and args.process and args.output_xls and args.thumbnails:
        baselight_contents = read_file(args.baselight)
        baselight_data = process_baselight_data(baselight_contents)
        xytech_contents = read_file(args.xytech)
        xytech_data, header_info, notes = process_xytech_data(xytech_contents)

        if args.process:
            video_duration, total_frames, fps = process_video(args.process)
            data_for_export = find_records_in_range(db, total_frames, fps, xytech_data, args.process, args.thumbnails)
            if data_for_export:
                export_to_xls(data_for_export, header_info, notes, args.output_xls, args.thumbnails)
                render_video_clips(data_for_export, args.process, "renders", fps)
            print("Data processed and exported to XLS.")
    
    upload_to_frameio(api_token, project_id, renders_folder)

if __name__ == "__main__":
    main()