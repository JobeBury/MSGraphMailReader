import asyncio
import configparser
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from graph import Graph
import os
import base64


async def main():
    print('Python Graph Tutorial\n')

    # Load settings
    config = configparser.ConfigParser()
    config.read(['config.cfg', 'config.dev.cfg'])
    azure_settings = config['azure']

    graph: Graph = Graph(azure_settings)

    await greet_user(graph)

    choice = -1

    while choice != 0:
        print('Please choose one of the following options:')
        print('0. Exit')
        print('1. Display access token')
        print('2. List my inbox')
        print('3. Send mail')
        print('4. List users')
        print('5. Get attachments')

        try:
            choice = int(input())
        except ValueError:
            choice = -1

        try:
            if choice == 0:
                print('Goodbye...')
            elif choice == 1:
                await display_access_token(graph)
            elif choice == 2:
                await list_inbox(graph)
            elif choice == 3:
                await send_mail(graph)
            elif choice == 4:
                await list_users(graph)
            elif choice == 5:
                await attachment_save(graph)
            elif choice == 6:
                await List_inbox(graph)
            else:
                print('Invalid choice!\n')
        except ODataError as odata_error:
            print('Error:')
            if odata_error.error:
                print(odata_error.error.code, odata_error.error.message)


async def greet_user(graph: Graph):
    CurrentUser = await graph.get_user()
    if CurrentUser:
        print('Hello,', CurrentUser.display_name)
        # For Work/school accounts, email is in mail property
        # Personal accounts, email is in userPrincipalName
        print('Email:', CurrentUser.mail or CurrentUser.user_principal_name, '\n')


async def display_access_token(graph: Graph):
    token = await graph.get_user_token()
    print('User token:', token, '\n')


async def list_inbox(graph: Graph):
    message_page = await graph.get_inbox()
    if message_page and message_page.value:
        # Output each message's details
        for message in message_page.value:
            print('Message:', message.subject)
            if (
                    message.from_ and
                    message.from_.email_address
            ):
                print('  From:', message.from_.email_address.name or 'NONE')
            else:
                print('  From: NONE')
            print('  Status:', 'Read' if message.is_read else 'Unread')
            print('  Received:', message.received_date_time)
            print(' Attachments', message.has_attachments)
        # If @odata.nextLink is present
        more_available = message_page.odata_next_link is not None
        print('\nMore messages available?', more_available, '\n')


async def send_mail(graph: Graph):
    # Send mail to the signed-in user
    # Get the user for their email address
    user = await graph.get_user()
    if user:
        user_email = user.mail or user.user_principal_name

        await graph.send_mail('Testing Microsoft Graph', 'Hello world!', user_email or '')
        print('Mail sent.\n')


async def list_users(graph: Graph):
    users_page = await graph.get_users()
    # Output each users' details
    if users_page and users_page.value:
        for user in users_page.value:
            print('User:', user.display_name)
            print('ID:', user.id)
            print('Email:', user.mail)
            print('department:', user.department)
            print('Location:', user.office_location)

        # If @odata.nextLink is present
        more_available = users_page.odata_next_link is not None
        print('\nMore users available?', more_available, '\n')


async def attachment_save(graph: Graph):
    message_page = await graph.get_inbox()

    if message_page and message_page.value:
        for message in message_page.value:
            # check if not read
            if not message.is_read:
                # check for attachments
                if message.has_attachments:
                    print(message.subject, 'has attachments!')

                    # Get the attachments for the current message
                    attachments_response = await graph.get_attachments(message.id)
                    attachments = attachments_response.value  # Access the 'value' property

                    if attachments:
                        for attachment in attachments:
                            if ".pdf" in attachment.name.lower():
                                attachment_name = attachment.name
                                attachment_id = attachment.id
                                attachment_content_type = attachment.content_type

                                # Retrieve the attachment content (basics)
                                print('Attachment Name:', attachment_name)
                                print('Attachment ID:', attachment_id)
                                print('Attachment Type:', attachment_content_type)
                                # Retrieve the attachment content (FileAttachment)
                                if "application/pdf" in attachment.content_type:
                                    print(message.subject, "contains a True PDF")
                                    attachment_content = await graph.get_attachment_content(message.id, attachment.id)
                                    print('Attachment bytes:', attachment_content)

                                    # decode the content

                                    attachment_content_decoded = base64.b64decode(attachment_content.content_bytes)

                                    # direct to directory
                                    save_path = r'C:\Users\JobeB\Desktop\MSGDownloads'
                                    file_name = attachment.name
                                    file_path = os.path.join(save_path, file_name)

                                    # Check if the file already exists
                                    counter = 1
                                    while os.path.exists(file_path):
                                        # If the file already exists
                                        file_name_without_extension, file_extension = os.path.splitext(attachment.name)
                                        new_file_name = f"{file_name_without_extension}_{counter}{file_extension}"
                                        file_path = os.path.join(save_path, new_file_name)
                                        print(file_path)
                                        counter += 1

                                    # Save the attachment
                                    print("Saving to File Path:", file_path)
                                    with open(file_path, 'wb') as file:
                                        file.write(bytes(attachment_content_decoded))
                                    print(f"Attachment saved successfully at: {file_path}")
                                    await graph.move_mail(message_id=message.id)
                                else:
                                    print(message.subject, "'s Attachment is not a PDF")
                                    await graph.move_mail(message_id=message.id)
                            else:
                                print(message.subject, "'s Attachment is not a PDF")
                                await graph.move_mail(message_id=message.id)
                    else:
                        print(message.subject, 'has no attachments')
                        await graph.move_mail(message_id=message.id)
                else:
                    print(message.subject, "has no attachments")
                    await graph.move_mail(message_id=message.id)
            else:
                print(message.subject, "has already been read")
                await graph.move_mail(message_id=message.id)


async def List_inbox(graph: Graph):
    inbox_list = await graph.List_inboxes()
    Inbox_list = inbox_list.value
    print(Inbox_list)

# Run main
asyncio.run(main())
