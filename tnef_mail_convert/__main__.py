from message import Message


fp = open("../testdata/rtf-attachments.eml", "rb")

message = Message()
message.parse_file(fp)

if message.has_winmail():
    print("This is a TNEF message")
    if len(message.new_attachments) > 0:
        print()
        print("Added new attachments:\n")
        for name in message.new_attachments:
            print("  " + name)

    with open("../testdata/output.eml", "wb") as fh:
        fh.write(message.get_message_without_winmail().as_bytes())

else:
    print("No TNEF data found!")
