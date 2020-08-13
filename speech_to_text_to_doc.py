import speech_recognition as sr
import  docx

r = sr.Recognizer()

with sr.Microphone() as source:
    print('Speak Anything : ')
    r.adjust_for_ambient_noise(source)
    audio = r.listen(source)
    
    try:
        document_name = input('Document name with extension: ')
        doc = docx.Document(document_name)
        text = r.recognize_google(audio)
        print('You said: {}'.format(text))
        doc.add_paragraph(text)
        doc.save(document_name)
        
    except:
        print('Sorry could not recognize your voice')