import poplib
from email import message_from_bytes
from email.header import decode_header


pop_server = 'outlook.office365.com'
pop_port = 995

#connect to the POP3 server and get welcome
pop_conn = poplib.POP3_SSL(pop_server, pop_port)

server_response = pop_conn.getwelcome().decode()

start_index = server_response.find('+')
end_index = server_response.find('.')

cleaned_response = server_response[start_index:end_index + 1]

#print the parsed greeting
print(cleaned_response)

#ask for username and send to server
email_address = input('Enter your email address: ')
response = pop_conn.user(email_address)
print(response.decode())

#ask for password and send to server
password_pop = input('Enter Password: ')
response = pop_conn.pass_(password_pop)
print(response.decode())

#function to handle STAT
def handle_stat():
    response = pop_conn.stat()
    num_messages = response[0]
    total_size = response[1]
    print(f"+OK {num_messages} {total_size}")
    
    
marked_for_deletion = []

#function to handle DELE
def handle_dele(index):
    response = pop_conn.list()
    message_list = response[1]
    message_indexes = [int(msg.decode().split()[0]) for msg in message_list]
    
    #checks whether message marked for deletion already
    
    if index in message_indexes:
        marked_for_deletion.append(index)
        print(f"-ERR Message {index} marked for deletion.")
    else:
        print(f"-ERR Message with index {index} does not exist.")
        
#function to handle LIST
def handle_list(index=None):
    
    #case when LIST is with an index
    if index is not None:
        try:
            response = pop_conn.list(index)
            response_str = response.decode()
            response_parts = response_str.split()
            message_index = response_parts[1]
            message_size = response_parts[2]
            print(f"+OK {message_index} {message_size} octets")
        except poplib.error_proto as e:
            print(f"{e}")

    #case when LIST is without index
    else:
        response = pop_conn.list()
        message_list = response[1]
        total_octets = sum(int(msg.decode().split()[1]) for msg in message_list)
        total_messages = len(message_list)
        print(f"+OK {total_messages} messages {total_octets} octets")
        for msg in message_list:
            print(msg.decode())

#function to handle QUIT
def handle_quit():
    #checks whether messages are marked for deletion, deletes them if so
    for index in marked_for_deletion:
        response = pop_conn.dele(index)
        status = response.decode()
        if status.startswith('+OK'):
            print(f"+OK Message {index} deleted.")
        else:
            print(f"-ERR Failed to delete message {index}.")
    pop_conn.quit()
    
#handles RSET
def handle_rset():
    response = pop_conn.stat()
    num_messages = len(marked_for_deletion)
    print(f"+OK {num_messages} messages unmarked from deletion")
    marked_for_deletion.clear()
    
#handles RETR
def handle_retr(index):
    try:
        response = pop_conn.retr(index)
        status = response[0].decode()
        message_content = b'\n'.join(response[1])

    #parses extracted message
        msg = message_from_bytes(message_content)
        sender = msg.get('From')
        recipient = msg.get('To')
        subject_bytes = msg.get('Subject')
        subject = decode_header(subject_bytes)[0][0]
        if isinstance(subject, bytes):
            subject = subject.decode()
        date_time = msg.get('Date')
        content_payload = None

    #gets message content
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

        #extract text content
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    content_payload = part.get_payload(decode=True)
                    if content_payload is not None:
                        charset = part.get_content_charset() or 'utf-8'  
                        content = content_payload.decode(charset, 'ignore')
                    break
        else:
            content_payload = msg.get_payload(decode=True)
            if content_payload is not None:
                charset = msg.get_content_charset() or 'utf-8'  
                content = content_payload.decode(charset, 'ignore')
    

#check if content_payload is not None
        if content_payload is not None:
            charset = msg.get_content_charset() or 'utf-8'  
            content = content_payload.decode(charset, 'ignore')

        print("Sender:", sender)
        print("Recipient:", recipient)
        print("Subject:", subject)
        print("Date/Time:", date_time)
        print("Message content:", content)
    except poplib.error_proto as e:
            print(f"{e}")

def handle_retr_all():
        responseR = pop_conn.list()
        message_list = responseR[1]

        for msg_info in message_list:
            msg_number, _ = msg_info.split()
            cleaned_numberT= int(msg_number)
            handle_retr(cleaned_numberT)
  

#main logic loop
while True:
    command = input().strip()
    if not command:
        print("-ERR Not valid command")
        continue
    elif command == 'STAT':
        handle_stat()
    elif command == 'NOOP':
        print(f"+OK")
    elif command.startswith('LIST'):
        try:
            index = int(command.split()[1])
            if index in marked_for_deletion:
                print(f"-ERR Message {index} is marked for deletion.")
            else:
                handle_list(index)
        except IndexError:
            handle_list()
        except ValueError:
            print("-ERR Invalid index.")
    elif command.startswith('DELE'):
        try:
            index = int(command.split()[1])
            if index in marked_for_deletion:
                print(f"-ERR Message {index} is marked for deletion.")
            else:
                handle_dele(index)
        except poplib.error_proto as e:
            print(f"Error: {e}")
        except ValueError:
            print("-ERR Invalid index.")
    elif command.startswith('RSET'):
        handle_rset()
    elif command.startswith('RETR'):
        try:
            index = int(command.split()[1])
            if index in marked_for_deletion:
                print(f"-ERR Message {index} is marked for deletion.")
            else:
                handle_retr(index)
        except ValueError:
            print("-ERR Invalid index.")
    elif command.startswith('GET ALL'):
        handle_retr_all()
    elif command == 'QUIT':
        handle_quit()
        print("+OK Bye-bye!")
        break
    else:
        print("-ERR Not a valid command")