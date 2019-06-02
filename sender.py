# -+- coding=utf8 -+-
import copy
import logging
import smtplib
import sys
import time
import random
from email.utils import COMMASPACE
import os
import sys

import time
import infoGetter
log_name = 'mail'
default_sender = 'dian@hust.edu.cn'
default_host = 'mail.hust.edu.cn'
default_port = 25
mails_per_send = 50
mails_per_login = 300
interval_between_login = 1
reset_per_n_login = 4
reset_interval = 5


class Server():
    def __init__(self, username: str = default_sender, password: str = 'password', host: str = default_host, port: int = default_port, debug: bool = True, ts=time.strftime('%Y-%m-%d-%H-%M-%S')):
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.debug = debug
        self.ts = ts

        self.fail_send = []
        self.all_send = []
        self.total_send = 0
        self.contierror = 0

        self._login = False

        self.logger = logging.getLogger(log_name)
        self.logger.setLevel(logging.DEBUG)
        fmt = logging.Formatter('[%(levelname)s] %(message)s')

        fh = logging.FileHandler(
            './logs/mail_{}.log'.format(self.ts), encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(fmt)
        self.logger.addHandler(fh)

        self.logger.debug(self.__repr__())

    def log_error(self):
        if self.fail_send:
            self.logger.critical('fail send!!!')
            if self.fail_send:
                self.logger.critical(' '.join(self.fail_send))

    def save_last_successful_send(self):
        with open(time.strftime('./logs/last_send_{}.txt'.format(self.ts)), 'w') as f:
            f.write(
                ' '.join([x for x in self.all_send if not x in self.fail_send]))

    def login(self):
        """Login"""
        if self.debug:
            self._login = True
            self.logger.debug('Login successful (DEBUG MODE)')
            return
        try:
            self.server = smtplib.SMTP(self.host, self.port)
        except:
            self.logger.error('can not connect to server')
            self.log_error()
            self.save_last_successful_send()
            sys.exit(3)

        if self._login:
            self.logger.warning('{} duplicate login!'.format(self.__repr__()))
            return

        try:
            self.server.login(self.username, self.password)
        except smtplib.SMTPAuthenticationError:
            self.logger.error('username or password error')
            self.log_error()
            self.save_last_successful_send()
            sys.exit(4)

        self._login = True
        self.logger.debug('Login successful')

    def logout(self):
        """Logout"""
        if self.debug:
            self._login = False
            self._needstls = True
            self.server = None
            self.logger.debug('Logout successful (DEBUG MODE)')
            return

        if not self._login:
            self.logger.warning(
                '{} Logout before login!'.format(self.__repr__()))
            return

        try:
            code, message = self.server.docmd("QUIT")
            if code != 221:
                raise smtplib.SMTPResponseException(code, message)
        except smtplib.SMTPServerDisconnected:
            pass
        finally:
            self.server.close()

        self.server = None
        self._login = False
        self.logger.debug('Logout successful')

    def check_available(self) -> bool:
        """test server"""
        try:
            self.login()
            self.logout()
            return True
        except Exception as e:
            self.logger.error('server does not available'+e)
            return False

    def is_login(self) -> bool:
        return self._login

    def __repr__(self):
        return '<{} username:{} password:{} {}:{}>'.format(self.__class__.__name__, self.username, self.password, self.host, self.port)

    def _send_mails(self, reciver, msg):
        if not self.is_login():
            self.login()
        msg['Bcc'] = COMMASPACE.join(reciver)
        msg = msg.as_string()
        self.logger.info(('Send to %d people ' %
                          len(reciver))+(' '.join(reciver)))
        ret = None
        try:
            if self.debug:
                self.logger.info('(DEBUG MODE)Send to %d' % len(reciver))
            else:
                ret = self.server.sendmail(
                    self.username, reciver, msg)
            self.contierror = 0
        except smtplib.SMTPSenderRefused as e:
            self.logger.error(e.__str__())
            self.fail_send.extend(reciver)
            self.log_error()
            self.save_last_successful_send()
            sys.exit(5)
        except smtplib.SMTPRecipientsRefused as e:
            self.logger.error(e.__str__())
            self.fail_send.extend(reciver)
            self.contierror += 1
            if self.contierror == 3:
                self.log_error()
                self.save_last_successful_send()
                sys.exit(6)
        except smtplib.SMTPDataError as e:
            self.logger.error(e.__str__())
            self.fail_send.extend(reciver)
            self.contierror += 1
            if self.contierror == 3:
                self.log_error()
                self.save_last_successful_send()
                sys.exit(7)

        self.all_send.extend(reciver)
        self.total_send += len(reciver)
        if ret:
            self.fail_send.extend(ret.keys())
        self.logger.info('==========Try======== %s' % self.total_send)

    def send_all_mails(self, reciver, mail):
        last = len(reciver)
        msg = mail.make_mail()
        if not msg:
            self.logger.error('msg error')
            sys.exit(8)
        self.logger.info('Sender: '+mail.nickname+' '+mail.addr)
        self.logger.info('Subject: '+mail.subject)
        self.logger.info('Reciver: {}'.format(last))
        self.logger.info('Content: \n'+mail.content_txt)
        self.logger.info('Attachment: '+(' '.join(mail.attachmentname)))
        self.logger.info('Debug Mode: ' + 'ON' if self.debug else 'OFF')
        if not self.is_login():
            self.login()
        send_index = 0
        counter = 0
        login_cnt = 1
        while last > 0:
            if last >= mails_per_send:
                self._send_mails(
                    reciver[send_index:send_index+mails_per_send], copy.deepcopy(msg))
                send_index += mails_per_send
                counter += mails_per_send
                last -= mails_per_send
            elif last > 0:
                self._send_mails(reciver[send_index:], copy.deepcopy(msg))
                send_index += last
                counter += last
                last = 0

            if last > 0 and counter+mails_per_send > mails_per_login:
                self.logger.info('Relogin')
                self.logout()
                if last > 0 and login_cnt > reset_per_n_login:
                    login_cnt = 0
                    self.logger.info('Interval')
                    if not self.debug:
                        time.sleep(reset_interval)
                else:
                    if not self.debug:
                        time.sleep(interval_between_login)
                self.login()
                counter = 0
                login_cnt += 1
        self.logout()
        self.logger.info('Finish send')
        self.log_error()
        self.save_last_successful_send()


if __name__ == "__main__":
    ts = sys.argv[-1]
    try:
        emails, passwd, *_, mail, dbg = infoGetter.get('./uploads/'+ts+'.zip')
    except Exception:
        sys.exit(1)
    # with open('temp.txt', 'w') as f:
    #     print(mail.make_mail().as_string(), file=f)
    # print(emails, mail, dbg)
    server = Server(password=passwd, debug=dbg, ts=ts)
    if not server.check_available():
        sys.exit(2)
    server.send_all_mails(emails, mail)
    sys.stdout.write('Finish\n')
    # time.sleep(30)
