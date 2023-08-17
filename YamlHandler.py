import yaml
import os

class YamlHandler:
    def __init__(self, file=None):
        self.file = os.path.join(self.GetDirName(), 'config.yaml') if file is None else file
        self.config = self.ReadYaml()

    def ReadYaml(self, encoding='utf-8'):
        """读取yaml数据"""
        with open(self.file, encoding=encoding) as f:
            return yaml.load(f.read(), Loader=yaml.FullLoader)

    def WriteYaml(self, data, encoding='utf-8'):
        """向yaml文件写入数据"""
        with open(self.file, encoding=encoding, mode='w') as f:
            return yaml.dump(data, stream=f, allow_unicode=True)

    def GetDirName(self):
        scriptPath = os.path.abspath(__file__)
        dirName = os.path.dirname(scriptPath)
        return dirName

Settings = YamlHandler().config
# yaml_data = YamlHandler(os.path.join(os.getcwd(), 'config.yaml')).ReadYaml()
