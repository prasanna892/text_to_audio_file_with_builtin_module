import win32com.client as wincl
import os

outfile = input("enter_file_name.extinsion: ")
text = input("Enter your text here: ")

speaker_number = 2  # speaker number
spk = wincl.Dispatch("SAPI.SpVoice")
filestream = wincl.Dispatch("SAPI.SpFileStream")
filestream.open(outfile, 3, False)
vcs = spk.GetVoices()
#print(vcs.Item(speaker_number).GetAttribute("Name")) # speaker name
spk.Voice
spk.Rate = 1.2  # voice speed rate
spk.SetVoice(vcs.Item(speaker_number)) # set voice (see Windows Text-to-Speech settings)
spk.AudioOutputStream = filestream
spk.Speak(text)
filestream.close()

print("Conversion completed")

os.system(outfile) # playing audio file