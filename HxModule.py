import sys
class HxModule:
    def __init__(mod_name,path=""):
        self.mod_name = mod_name

        if path not in ["",None]:
            sys.path.append(path)

    def unload(self):
        del(self.mod)
        del(sys.modules[self.mod_name])

    def load(self):
        self.mod = __import__(self.mod_name)

    def reload(self):
        self.unload()
        self.load()
