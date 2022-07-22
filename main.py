import speech_recognition as sr
import aspose.words as aw
from docx import Document

r = sr.Recognizer()
# create document object
doc = aw.Document()

def listen_voice():
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            voice_data = r.recognize_google(audio)
            return voice_data
        except sr.UnknownValueError:
            print('I did not quite catch it')
        except sr.RequestError:
            print('Sorry the servers are down')

def write_file(voice_data):
    num = 1

    # create file name
    t_file_name = 'text' + str(num) + '.txt'
    w_file_name = 'word' + str(num) + '.docx'
    wd_file_name = 'word_doc' + str(num) + '.docx'

    # create text file
    fp = open(t_file_name, 'a')

    # create word file
    # create a document builder object
    builder = aw.DocumentBuilder(doc)

    # create word using document
    document = Document()

    # write voice data into file
    if 'document is complete' not in voice_data:
        print(voice_data)

        # wite to text
        fp.write(voice_data + '\n')

        # write to word
        builder.writeln(voice_data)

        # write paragraph using document
        # document.add_paragraph("voice_data")
        paragraph = document.add_paragraph('Lorem ipsum dolor sit amet.')

        print('written, please continue')
    else:
        print('closing document')

        # save word document
        doc.save(w_file_name)

        # close text document
        fp.close()

        # save using document
        document.save(wd_file_name)

        num += 1
        exit()

print('let us create a document... listening')
while(1):
    voice_data = listen_voice()
    write_file(voice_data)