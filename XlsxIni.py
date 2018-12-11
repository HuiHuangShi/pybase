from configparser import ConfigParser
from NameTuple import NameTuple
import re

#
#[core]
#    column_range = [0,3]
#    value = { var:"xxx"} notice:if excel the first line have '\n',it will remove it
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

    def __is_key_word(self,word):
        for check_item in self.check_key_lst:
            if word.find(check_item) > 0:
                return False
        return True

    def __parse_core(self):
        #parse column_range
        self.key_range = self.__core_column_range_regex.findall(self.core.column_range)[0]
        if "" in self.key_range or None in self.key_range:
            raise Exception("column_range:%s format is incorrect" % self.key_range)

        if self.key_range[0] == '(':
            start = int(self.key_range[1]) + 1
        else:
            start = int(self.key_range[1])

        if self.key_range[-1] == ')':
            end = int(self.key_range[-2]) - 1
        else:
            end = int(self.key_range[-2])

        self.column_range = [start, end]

        #parse key_list
        if self.core.key_list[0] != '[' or self.core.key_list[-1] != ']':
            raise Exception("key_list:%s format is incorrect" % self.core.key_list)
        key_lst = self.core.key_list[1:-1].split(",")

        for key in key_lst:
            if key in [None,""]:
                raise Exception("key_list:%s format is incorrect" % key_lst)

            if not self.__is_key_word(key):
                raise Exception("key_list:%s format is incorrect" % key)

        #check key_list len is in the range of column_range
        length = self.column_range[1] - self.column_range[0] + 1
        if len(key_lst) != length:
            print("key_lst:%s" % key_lst)
            print("column_range:%s" % self.column_range)
            raise Exception("column_range and key_list is not match")

        i = 0
        for idx in xrange(self.column_range[0],self.column_range[1] + 1):
            self.range_map.append([idx,key_lst[i]])
            if key_lst[i] in self.struct.__dict__:
                raise Exception("already have name:%s" % key_lst[i])
            else:
                self.struct.__dict__[key_lst[i]] = idx
            i = i + 1

        #parse value_map
        value_map = ''.join(self.core.value_map.split("\n")).strip()
        map_pairs = value_map[1:-1].split(",")
        if "" in map_pairs or None in map_pairs:
            raise Exception("value_map content is not correct,please check")

        check_val_lst = [" ","\n"]
        for map_pair in map_pairs:
            res = map_pair.split(":")
            if len(res) != 2:
                raise Exception("%s is incorrect format" % map_pair)
            if res[0] in ["",None] or res[1] in ["",None]:
                raise Exception("%s is incorrect format" % map_pair)

            raw_map_key = res[0].strip()
            raw_map_val = res[1].strip()
            if not self.__is_key_word(raw_map_key):
                raise Exception("map_key:%s is incorrect format" % raw_map_key)

            for check_item in check_val_lst:
                if raw_map_val.find(check_item) > 0:
                    raise Exception("map_val:%s is incorrect format" % raw_map_val)
            if raw_map_val[0] != '"' or raw_map_val[-1] != '"':
                    raise Exception("map_val:%s is incorrect format" % raw_map_val)

            if raw_map_key in self.map.__dict__:
                raise Exception("%s is already in map" % raw_map_key)
            self.map.__dict__[raw_map_key] = raw_map_val

        #parse mandatory
        mandatory_lst = self.core.mandatory.split(",")
        for mandatory in mandatory_lst:
            if not self.__is_key_word(mandatory):
                raise Exception("mandatory:%s is incorrect format" % mandatory)
            if mandatory not in self.map.__dict__:
                raise Exception("mandatory:%s is not exist in value_map" % mandatory)
            if mandatory not in self.mandatory:
                self.mandatory.append(mandatory)
            else:
                raise Exception("mandatory:%s is repeated" % mandatory)

        #parse is_structure
        if self.core.is_structure not in ["yes","no"]:
            raise Exception("is_structure only have two pattern:yes/no")

    def __init__(self,ini_file):
        self.conf = ConfigParser()
        self.__core_column_range_regex = re.compile("^\s*(\(|\[)\s*(\d+)\s*(,)\s*(\d+)\s*(\)|\])")
        self.struct = NameTuple(["default"])
        self.map = NameTuple(["default"])
        self.column_range = []
        self.range_map = []
        self.mandatory = []
        self.check_key_lst = [" ","\n","\"","\'"]

        self.conf.read(ini_file)
        self.keys = self.conf.keys()

        if "core" not in self.keys:
            raise Exception("%s file not have core section")

        core_must_keys = ["column_range","key_list",
                               "value_map","mandatory","is_structure"]
        self.__assign_mandatory_attribute("core", core_must_keys, self.conf)
        self.__parse_core()
