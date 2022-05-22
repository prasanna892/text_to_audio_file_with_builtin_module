from comtypes.client import CreateObject    
from comtypes.gen import SpeechLib
import os

outfile = input("enter_file_name.extinsion: ")
text = input("Enter your text here: ")

engine = CreateObject("SAPI.SpVoice")
stream = CreateObject("SAPI.SpFileStream")
stream.Open(outfile, SpeechLib.SSFMCreateForWrite)
engine.AudioOutputStream = stream
engine.speak(text)
stream.Close()

print("Conversion completed")

os.system(outfile) # playing audio file