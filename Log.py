import logging
import pytz
import datetime
import os


class Log:
    def __init__(self, level='DEBUG'):
        self.log = logging.getLogger('KK')
        self.log.setLevel(level)
        self.beiJingTimeZone = pytz.timezone('Asia/Shanghai')
        self.logPath = self.GetLogPath()

    def ConsoleHandle(self, level='DEBUG'):
        '''控制台处理器'''
        consoleHandler = logging.StreamHandler()
        consoleHandler.setLevel(level)
        consoleHandler.setFormatter(self.GetFormatter()[0])
        consoleHandler.formatter.converter = lambda *args: datetime.datetime.now(self.beiJingTimeZone).timetuple()
        return consoleHandler

    def FileHandle(self, level='DEBUG'):
        '''文件处理器'''
        fileHandler = logging.FileHandler(self.logPath, mode='a', encoding='utf-8')
        fileHandler.setLevel(level)
        fileHandler.setFormatter(self.GetFormatter()[1])
        fileHandler.formatter.converter = lambda *args: datetime.datetime.now(self.beiJingTimeZone).timetuple()
        return fileHandler

    def GetFormatter(self):
        '''格式器'''
        consoleFmt = logging.Formatter(fmt='[%(asctime)s %(levelname)s]:%(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')
        fileFmt = logging.Formatter(fmt='[%(asctime)s %(levelname)s]:%(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')
        return consoleFmt, fileFmt

    def GetLog(self):
        '''日志器添加到控制台处理器'''
        if not self.log.handlers:
            self.log.addHandler(self.ConsoleHandle())
            self.log.addHandler(self.FileHandle())
        return self.log

    def GetDirName(self):
        scriptPath = os.path.abspath(__file__)
        dirName = os.path.dirname(scriptPath)
        return dirName
    
    def GetLogPath(self):
        dirName = self.GetDirName()
        logPath = os.path.join(dirName, 'log.txt')
        return logPath
