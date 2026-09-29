Put golden-test pairs here (they are git-ignored, since teacher copies contain answer keys):

    python3 test/docx_to_json.py "<teacher>.docx" > test/fixtures/<name>.teacher.json
    python3 test/docx_to_json.py "<student>.docx" > test/fixtures/<name>.student.json

`npm test` then checks that converting the teacher copy gives the student copy.
To get a .docx from a Google Doc: File > Download > Microsoft Word.
