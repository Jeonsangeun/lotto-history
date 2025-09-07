from pytube import YouTube
import os
import ffmpeg

PATH = "https://www.youtube.com/watch?v=cnNN-KryjbU"
folder_name = os.getcwd()
print(folder_name)
yt = YouTube(PATH)
'''
[<Stream: itag=”137″ mime_type=”video/mp4″ res=”1080p” fps=”30fps” vcodec=”avc1.640028″ progressive=”False” type=”video”>,
<Stream: itag=”248″ mime_type=”video/webm” res=”1080p” fps=”30fps” vcodec=”vp9″ progressive=”False” type=”video”>,
<Stream: itag=”399″ mime_type=”video/mp4″ res=”None” fps=”30fps” vcodec=”av01.0.08M.08″ progressive=”False” type=”video”>,
<Stream: itag=”250″ mime_type=”audio/webm” abr=”70kbps” acodec=”opus” progressive=”False” type=”audio”>,
<Stream: itag=”251″ mime_type=”audio/webm” abr=”160kbps” acodec=”opus” progressive=”False” type=”audio”>]
'''
result= yt.streams.filter(file_extension='mp4')

for stream in result:
    print(stream)

stream_result = yt.streams.get_by_itag(137).download()


# Audio
# # audio_stream = yt.streams.filter(only_audio=True).desc().first()
# print(audio_stream)
# # 128kbps audio download
# # audio_stream = yt.streams.get_by_itag(140)
#
# # yt.streams.filter(file_extension='mp4')
# # stream_video = yt.streams.get_by_itag(137)
#
# # download
# audio_stream.download(folder_name)
#
# # file name
# file_name = audio_stream.default_filename
# base, ext = os.path.splitext(file_name)
# new_file = base + '.mp3'
# print(new_file)
#
# os.rename(file_name, new_file)
# # mp4 -> mp3

# ffmpeg.input(os.path.join(folder_name, file_name)).output(os.path.join(folder_name, new_file)).run(overwrite_output=True)
# os.remove(os.path.join(folder_name, file_name))

