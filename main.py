#!/usr/bin/env python
"""
Module for extracting all attachments from an mbox archive.

rework of https://github.com/treymo/google-takeout-helper
"""
 
import argparse   
import sys
import mailbox
import os
import re
import time
import datetime
import email
import pathlib
from email.policy import default


class MboxReader:
    """ lifted from https://stackoverflow.com/questions/59681461/read-a-big-mbox-file-with-python. """
    def __init__(self, filename):
        self.handle = open(filename, 'rb')
        assert self.handle.readline().startswith(b'From ')

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, exc_traceback):
        self.handle.close()

    def __iter__(self):
        return iter(self.__next__())

    def __next__(self):
        lines = []
        while True:
            line = self.handle.readline()
            if line == b'' or line.startswith(b'From '):
                yield email.message_from_bytes(b''.join(lines), policy=default)
                if line == b'':
                    break
                lines = []
                continue
            lines.append(line)


class Message:
    def __init__(self, message):
        self.message = message
    
    @property
    def subject(self):
        return self.message.get('subject', '')
    
    def is_spam(self, ignore_labels=['Spam']):
        for label in ignore_labels:
             if label in self.message.get('X-Gmail-Labels', ''):
                return True
        return False
    
    @property
    def has_attachment(self):
        return self.message.is_multipart()

    @property
    def sender(self):
        _sender = self.message.get('From', '')
        return _sender.split('<')[-1].split('>')[0]

    @property
    def sent(self):
        try:
            date_string = self.message['Date']
            sent = datetime.datetime.strptime(date_string, '%a, %d %b %Y %H:%M:%S %z')
            return sent.strftime('%Y-%b')
        except Exception:
            pass

    def iter_attachments(self):
        return self.message.iter_attachments()

    def __str__(self):
        return f'{self.sent} {self.sender:<50} {self.subject}'


def strip_illegal_char(input_string):
    return re.sub(r'[\<,\>,\:,\",\/,\\,\|,\?,\*,\n,\t]', '', str(input_string))


def extract_mail_attachments(mbox_file_path, ignore_labels, output_directory):
    with MboxReader(mbox_file_path) as mbox:
        for message in mbox:
            message = Message(message)

            if message.is_spam(ignore_labels):
                continue

            if not message.has_attachment:
                continue
                
            if not message.sender or not message.subject or not message.sent:
                continue

            print(message)

            for attachment in message.iter_attachments():
                filename = attachment.get_filename()
                if not filename:
                    continue
                filename = strip_illegal_char(filename)
                d = os.path.join(output_directory, message.sent)
                p = os.path.join(d, f'{message.sent}--{message.sender}--{filename}')
                os.makedirs(d, exist_ok=True)
                with open(p, 'wb') as a:
                    try:
                        a.write(attachment.get_payload(decode=True))
                    except Exception as e:
                        continue
                print(f'    {p}')

def main():
    parser = argparse.ArgumentParser(description='Extract all the attachments from a mbox file.')
    parser.add_argument('-i', '--ignore-labels', metavar='label', type=str, nargs='*', default=['spam'])
    parser.add_argument('-o', '--output-directory', metavar='attachments', type=pathlib.Path, help='directory to store all the attachments found in your mbox', default='attachments')
    parser.add_argument('mbox_file_path', metavar='/path/to/your/mail.mbox', type=pathlib.Path, help='path to your mbox file')
    args = parser.parse_args()
    print(args)
    extract_mail_attachments(args.mbox_file_path, args.ignore_labels, args.output_directory)


if __name__ == '__main__':
    main()
