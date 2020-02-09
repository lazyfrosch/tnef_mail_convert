"""
Message conversion from MIME with winmail.dat to extracted attachments
"""

import base64
from email.parser import Parser
import email.message as message
from tnefparse import TNEF
from typing import Union

WINMAIL_NAME = 'winmail.dat'
WINMAIL_RENAMED = 'original-winmail.dat'


class Message(object):
    tnef_message: Union[TNEF, None]
    tnef_payload: Union[message.Message, None]

    def __init__(self):
        self.original_body = None
        self.message = message.Message()
        self.tnef_message = None
        self.tnef_payload = None
        self.new_attachments = []

    def parse(self, text: str):
        parser = Parser()
        self.original_body = text
        self.message = parser.parsestr(text)
        self._read_tnef()
        if self.has_winmail():
            self._extract_htmlbody()
            self._extract_rtfbody()
            self._extract_attachments()
            self._rename_winmail()

    def has_winmail(self):
        if not self.message.is_multipart():
            return False
        return self.tnef_message is not None

    def __str__(self):
        return self.message.__str__()

    def _read_tnef(self):
        # TODO: does this work in non-multipart?
        payloads = self.message.get_payload()
        if not isinstance(payloads, list):
            return

        for payload in payloads:
            if payload.get_content_type() == "application/ms-tnef":
                # TODO: skip renamed winmail.dat
                data = base64.b64decode(payload.get_payload())
                self.tnef_payload = payload
                self.tnef_message = TNEF(data)
                return

    def _extract_htmlbody(self):
        """
        Extract an HTML body in winmail to the normal HTML body
        TODO: check how this is coming from Outlook and how to correct
        """
        pass

    def _extract_rtfbody(self):
        """
        Extract an RTF body into a new attachment
        """
        rtf = self.tnef_message.rtfbody
        if rtf is None:
            return

        # extract RTF body to an attachment
        filename = "mail-body.rtf"
        new_payload = message.Message()
        new_payload.set_type("application/rtf")
        new_payload.add_header("Content-Disposition", "attachment", filename=filename)
        new_payload.add_header("Content-Transfer-Encoding", "base64")
        new_payload.set_payload(encode_payload(rtf))
        self.message.attach(new_payload)
        self.new_attachments.append(filename)

    def _extract_attachments(self):
        """
        Extract every attachment to a regular MIME attachment
        """
        for attachment in self.tnef_message.attachments:
            new_payload = message.Message()
            new_payload.set_type("application/octet-stream")
            new_payload.add_header("Content-Disposition", "attachment", filename=attachment.long_filename())
            new_payload.add_header("Content-Transfer-Encoding", "base64")
            new_payload.set_payload(encode_payload(attachment.data))
            self.message.attach(new_payload)
            self.new_attachments.append(attachment.long_filename())

    def _rename_winmail(self):
        """
        Replace meta data for winmail.dat, so the file can be recognized as replaced

        But we want to keep the original

        TODO: What about X-MS-TNEF-Correlator ?
        """
        if self.tnef_payload:
            self.tnef_payload.replace_header("Content-Type", 'application/octet-stream; name="%s"' % WINMAIL_RENAMED)
            self.tnef_payload.replace_header("Content-Disposition", 'attachment; filename="%s"' % WINMAIL_RENAMED)


def encode_payload(data):
    encoded = base64.encodebytes(data)
    return encoded.decode('utf-8')
