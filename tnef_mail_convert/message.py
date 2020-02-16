"""
Message conversion from MIME with winmail.dat to extracted attachments
"""

import base64
from email.parser import Parser
import email.message
import email
from tnefparse import TNEF
from typing import Union, IO
import copy

WINMAIL_NAME = 'winmail.dat'
WINMAIL_RENAMED = 'original-winmail.dat'


class Message(object):
    message: email.message.Message
    tnef_message: Union[TNEF, None]
    tnef_payload: Union[email.message.Message, None]

    def __init__(self):
        self.message = email.message.Message()
        self.tnef_message = None
        self.tnef_payload = None
        self.new_attachments = []

    def parse(self, text: str):
        parser = Parser()
        self.message = parser.parsestr(text)
        self._parse()

    def parse_file(self, fp: IO[bytes]):
        self.message = email.message_from_binary_file(fp)
        self._parse()

    def _parse(self):
        self._read_tnef()
        if self.has_winmail():
            self._extract_body()
            self._extract_htmlbody()
            self._extract_rtfbody()
            self._extract_attachments()

    def has_winmail(self):
        if not self.message.is_multipart():
            return False
        return self.tnef_message is not None

    def __str__(self):
        return self.as_string()

    def as_string(self):
        return self.message.as_string()

    def get_message_without_winmail(self):
        result = copy.copy(self.message)

        del result["X-MS-TNEF-Correlator"]

        payloads = []
        for payload in self.message.get_payload():
            if payload.get_content_type() != "application/ms-tnef":
                payloads.append(payload)

        result.set_payload(payloads)

        return result

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

    def _extract_body(self):
        """
        Extract an text body in winmail to the normal body
        """
        body = self.tnef_message.body
        if body is None:
            return

        new_payload = email.message.Message()
        new_payload.set_type("text/plain")
        new_payload.set_payload(body)
        self.message.attach(new_payload)

    def _extract_htmlbody(self):
        """
        Extract an HTML body in winmail to the normal HTML body
        """
        html = self.tnef_message.htmlbody
        if html is None:
            return

        new_payload = email.message.Message()
        new_payload.set_type("text/html")
        # TODO: do we need base64?
        new_payload.set_payload(html)
        self.message.attach(new_payload)

    def _extract_rtfbody(self):
        """
        Extract an RTF body into a new attachment
        """
        rtf = self.tnef_message.rtfbody
        if rtf is None:
            return

        # extract RTF body to an attachment
        filename = "mail-body.rtf"
        new_payload = email.message.Message()
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
            new_payload = email.message.Message()
            new_payload.set_type("application/octet-stream")
            new_payload.add_header("Content-Disposition", "attachment", filename=attachment.long_filename())
            new_payload.add_header("Content-Transfer-Encoding", "base64")
            new_payload.set_payload(encode_payload(attachment.data))
            self.message.attach(new_payload)
            self.new_attachments.append(attachment.long_filename())


def encode_payload(data):
    encoded = base64.encodebytes(data)
    return encoded.decode('utf-8')
