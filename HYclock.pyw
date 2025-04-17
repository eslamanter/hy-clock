import time
import ctypes
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import pygame
import threading

pygame.mixer.init()

def get_current_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    current_volume = volume.GetMasterVolumeLevel()
    return current_volume

def set_max_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMute(0, None)
    volume.SetMasterVolumeLevel(0.0, None)

def set_volume(level):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevel(level, None)

def play_audio(num_times, stop_event):
    for _ in range(num_times):
        if stop_event.is_set():
            break
        pygame.mixer.music.load('audio.wav')
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            if stop_event.is_set():
                pygame.mixer.music.stop()
                break
            pygame.time.Clock().tick(10)

def show_message_box(message):
    return ctypes.windll.user32.MessageBoxW(0, message, "HYclock", 1)

while True:
    current_hour = time.localtime().tm_hour
    num_times_to_play = current_hour
    original_volume = get_current_volume()
    set_max_volume()
    stop_event = threading.Event()
    audio_thread = threading.Thread(target=play_audio, args=(num_times_to_play, stop_event))
    audio_thread.start()
    result = show_message_box(f"Bomberissimo! Sono le {current_hour}.")
    if result == 2:
        stop_event.set()
    audio_thread.join()
    set_volume(original_volume)
    time.sleep(3600 - time.localtime().tm_min * 60 - time.localtime().tm_sec)
