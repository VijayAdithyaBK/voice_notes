import speech_recognition as sr

r = sr.Recognizer()

'''
listen to voice audio
convert it into text
'''
def listen_voice(ask = False):
    with sr.Microphone() as source:
        if(ask):
            print(ask)
        audio = r.listen(source)
        try:
            voice_data = r.recognize_google(audio)
            return voice_data
        except sr.UnknownValueError:
            print('I did not quite catch it')
        except sr.RequestError:
            print('Sorry the servers are down')

'''
create file with desired file name
returns file pointer
'''
def create_file(voice_data):
    # create file name
    file_name = 'documents/' + voice_data + '.docx'
    # create text file
    try:
        # creates new file only if it does not exist
        fp = open(file_name, 'a+')
    except FileExistsError:
        print('file already exists')
        ask()
    return fp

'''
write into document
takes voice data and file pointer as  arguments
'''
def write_file(voice_data, fp):
    # write voice data into file
    if 'document is complete' != voice_data:
        print(voice_data)
        if 'next line' == voice_data:
            fp.write('\n')
        elif 'period' == voice_data:
            fp.write('.')
        else:
            # wite to file
            fp.write(voice_data + ' ')

        print('written, please continue')
    elif 'document is complete' == voice_data:
        print('closing document')

        # close text document
        fp.close()

        print('create new document or exit?')
        voice_data = listen_voice()
        if 'new document' in voice_data:
            ask()
        else:
            exit()

def ask():
    print('let us create a document')
    voice_data = listen_voice('what is name of document')
    fp = create_file(voice_data)
    print('please begin the document')
    while(1):
        voice_data = listen_voice()
        write_file(voice_data, fp)

ask()