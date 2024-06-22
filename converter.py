from PIL import Image as i
import subprocess
from tkinter import messagebox as me
import os,sys

def subprocess_args(include_stdout=True):
    if hasattr(subprocess, 'STARTUPINFO'):

        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        env = os.environ
    else:
        si = None
        env = None

    if include_stdout:
        ret = {'stdout': subprocess.PIPE}
    else:
        ret = {}

    ret.update({'stdin': subprocess.PIPE,
                'stderr': subprocess.PIPE,
                'startupinfo': si,
                'env': env })
    return ret


def re(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(filename)

def image_convert(file_path:str,output_path:str):
    try:
        if output_path.endswith('ico'):
            convert = i.open(file_path).save(f'{output_path}',sizes=[(90,90)])
        else:
            convert = i.open(file_path).save(f'{output_path}')
        return convert
    except FileNotFoundError as e:
        return "FileNotFoundError"
    except FileExistsError as e:
        return "FileExistsError"
    except Exception as e:
        print(e)
        return False

def audio_convert(file_path:str,output_path:str):
    try:
        convert = subprocess.run(f'{re("resource\\ffmpeg.exe")} -c:v libopenh264 -y -i {file_path} {output_path}', **subprocess_args(True))
        return convert
    except FileNotFoundError as e:
        return "FileNotFoundError"
    except FileExistsError as e:
        return "FileExistsError"
    except Exception as e:
        print(e)
        return False

def video_convert(file_path:str,output_path:str):
    try:
        subprocess.run(f'{re("resource\\ffmpeg.exe")} -c:v libopenh264 -y -i {file_path} {output_path}', **subprocess_args(True))
        return True
    except FileNotFoundError as e:
        return "FileNotFoundError"
    except FileExistsError as e:
        return "FileExistsError"
    except Exception as e:
        print(e)
        return False
