from typing import List
import base64
from tnefparse import TNEF
from email.message import Message
from email.parser import Parser


def mail_payload_encode(bytes):
    b64 = base64.encodebytes(bytes)
    return b64.decode('utf-8')


contents = ""
with open("../testdata/rtf-attachments.eml", "r") as fh:
    contents += fh.read()

parser = Parser()
mail = parser.parsestr(contents)

if mail.is_multipart():
    payloads: List[Message] = mail.get_payload()

    for i, payload in enumerate(payloads):
        if payload.get_content_type() == "application/ms-tnef":
            bytes = base64.b64decode(payload.get_payload())
            t = TNEF(bytes)

            if t.htmlbody:
                # TODO: extract HTML body
                pass

            rtf = t.rtfbody
            if rtf is not None:
                # extract RTF body to an attachment
                print("Converting RTF body to an attachment")
                new_payload = Message()
                new_payload.set_type("application/rtf")
                new_payload.add_header("Content-Disposition", "attachment", filename="winmail-content.rtf")
                new_payload.add_header("Content-Transfer-Encoding", "base64")
                new_payload.set_payload(mail_payload_encode(rtf))
                mail.attach(new_payload)

            # list TNEF attachments
            print("Attachments:\n")
            for a in t.attachments:
                print("  Extract Attachment: " + a.long_filename())
                new_payload = Message()
                new_payload.set_type("application/octet-stream")
                new_payload.add_header("Content-Disposition", "attachment", filename=a.long_filename())
                new_payload.add_header("Content-Transfer-Encoding", "base64")
                new_payload.set_payload(mail_payload_encode(a.data))
                mail.attach(new_payload)

    with open("../testdata/output.eml", "w") as fh:
        fh.write(mail.__str__())
