from message import Message


contents = ""
with open("../testdata/rtf-attachments.eml", "r") as fh:
    contents += fh.read()

message = Message()
message.parse(contents)

if message.has_winmail():
    print("This is a TNEF message")
    if len(message.new_attachments) > 0:
        print()
        print("Added new attachments:\n")
        for name in message.new_attachments:
            print("  " + name)

    with open("../testdata/output.eml", "w") as fh:
        fh.write(message.__str__())

else:
    print("No TNEF data found!")
