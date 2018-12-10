from configparser import ConfigParser
from NameTuple import NameTuple
import re

#
#[core]
#    column_range = [0,3]
class XlsxIni:
    def __assign_mandatory_attribute(self,attr,must_keys,config):
        if attr in self.__dict__:
            raise Exception("%s attr is already set" % attr)
        self.__dict__[attr] = NameTuple(must_keys)

        for must in must_keys:
            if must not in config[attr].keys():
                raise Exception("%s is mandotary in %s section" % (must,attr))
            else:
                self.__dict__[attr].__dict__[must] = config[attr][must]

    def __parse_core(self):
        ret = self.__core_column_range_regex.findall(self.core.column_range)

    def __init__(self,ini_file):
        self.conf = ConfigParser()
        self.__core_column_range_regex = re.compile("^\s*(\(|\[)\s*(\d+)\s*(,)\s*(\d+)\s*(\)|\])")

        self.conf.read(ini_file)
        self.keys = self.conf.keys()

        if "core" not in self.keys:
            raise Exception("%s file not have core section")

        core_must_keys = ["column_range","value_list",
                               "map","mandatory","is_structure"]
        self.__assign_mandatory_attribute("core", core_must_keys, self.conf)
        self.__parse_core()
