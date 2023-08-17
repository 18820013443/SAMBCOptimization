import logging
import pytz
import datetime
import os
from YamlHandler import Settings


class Log:
    def __init__(self):
        # 读取yaml中logging配置
        config = Settings.get('logging')

        # 获取并配置logging对象
        self.log = logging.getLogger('KK')
        self.logLevel = config.get('logLevel')
        self.log.setLevel(self.logLevel)

        # logging的配置
        self.fmt = config.get('fmt')
        self.dateFmt = config.get('dateFmt')
        self.fileName = config.get('logFileName')
        self.timeZone = pytz.timezone(config.get('timeZone'))
        self.logPath = self.GetLogPath() if config.get('logPath') == '' else config.get('logPath')

    def _ConsoleHandle(self):
        '''控制台处理器'''
        consoleHandler = logging.StreamHandler()
        consoleHandler.setLevel(self.logLevel)
        consoleHandler.setFormatter(self.GetFormatter()[0])
        consoleHandler.formatter.converter = lambda *args: datetime.datetime.now(self.timeZone).timetuple()
        return consoleHandler

    def _FileHandle(self):
        '''文件处理器'''
        fileHandler = logging.FileHandler(self.logPath, mode='a', encoding='utf-8')
        fileHandler.setLevel(self.logLevel)
        fileHandler.setFormatter(self.GetFormatter()[1])
        fileHandler.formatter.converter = lambda *args: datetime.datetime.now(self.timeZone).timetuple()
        return fileHandler

    def GetFormatter(self):
        '''格式器'''
        # consoleFmt = logging.Formatter(fmt='[%(asctime)s %(levelname)s]:%(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')
        # fileFmt = logging.Formatter(fmt='[%(asctime)s %(levelname)s]:%(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')
        consoleFmt = logging.Formatter(self.fmt, self.dateFmt)
        fileFmt = logging.Formatter(self.fmt, self.dateFmt)
        return consoleFmt, fileFmt

    def GetLog(self):
        '''日志器添加到控制台处理器'''
        if not self.log.handlers:
            self.log.addHandler(self._ConsoleHandle())
            self.log.addHandler(self._FileHandle())
        return self.log

    def GetDirName(self):
        scriptPath = os.path.abspath(__file__)
        dirName = os.path.dirname(scriptPath)
        return dirName
    
    def GetLogPath(self):
        dirName = self.GetDirName()
        logPath = os.path.join(dirName, self.fileName)
        return logPath

logger = Log().GetLog()
    
if __name__ == '__main__':
    obj = Log()


