from configparser import ConfigParser
from NameTuple import NameTuple
import re

#
#[core]
#    column_range = [0,3]
#    value = { var:"xxx"} notice:if excel the first line have '\n',it will remove it
class XlsxIni:
    def __assign_mandatory_attribute(self,section,must_keys,config):
        if section in self.__dict__:
            raise Exception("%s section is already init" % section)
        self.__dict__[section] = NameTuple([""])
        self.__dict__[section].__dict__["raw"] = NameTuple(must_keys)
        this_raw = self.__dict__[section].__dict__["raw"]

        for must in must_keys:
            if must not in config[section].keys():
                raise Exception("%s is mandotary in %s section" % (must, section))
            else:
                this_raw.__dict__[must] = config[section][must]

    def __is_key_word(self,word):
        for check_item in self.check_key_lst:
            if word.find(check_item) > 0:
                return False
        return True

    def __parse_section(self,section):
        this_section = self.__dict__[section]
        this_raw = self.__dict__[section].__dict__["raw"]

        this_section.__dict__["column_range"] = None
        this_section.__dict__["range_map"] = []
        this_section.__dict__["struct"] = NameTuple([""])
        this_section.__dict__["map"] = NameTuple([""])
        this_section.__dict__["mandatory"] = []
        this_section.__dict__["is_struct"] = False

        #parse column_range
        this_raw.key_range = self.__column_range_regex.findall(this_raw.column_range)[0]
        if "" in this_raw.key_range or None in this_raw.key_range:
            raise Exception("column_range:%s format is incorrect" % this_raw.key_range)

        if this_raw.key_range[0] == '(':
            start = int(this_raw.key_range[1]) + 1
        else:
            start = int(this_raw.key_range[1])

        if this_raw.key_range[-1] == ')':
            end = int(this_raw.key_range[-2]) - 1
        else:
            end = int(this_raw.key_range[-2])

        this_section.column_range = [start, end]

        #parse key_list
        if this_raw.key_list[0] != '[' or this_raw.key_list[-1] != ']':
            raise Exception("key_list:%s format is incorrect" % thie_raw.key_list)
        key_lst = this_raw.key_list[1:-1].split(",")

        for key in key_lst:
            if key in [None,""]:
                raise Exception("key_list:%s format is incorrect" % key_lst)

            if not self.__is_key_word(key):
                raise Exception("key_list:%s format is incorrect" % key)

        #check key_list len is in the range of column_range
        length = this_section.column_range[1] - this_section.column_range[0] + 1
        if len(key_lst) != length:
            print("key_lst:%s" % key_lst)
            print("column_range:%s" % self.column_range)
            raise Exception("column_range and key_list is not match")

        i = 0
        for idx in xrange(this_section.column_range[0],this_section.column_range[1] + 1):
            this_section.range_map.append([idx, key_lst[i]])
            if key_lst[i] in this_section.struct.__dict__:
                raise Exception("already have name:%s" % key_lst[i])
            else:
                this_section.struct.__dict__[key_lst[i]] = idx
            i = i + 1

        #parse value_map
        value_map = ''.join(this_raw.value_map.split("\n")).strip()
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

            if raw_map_key in this_section.map.__dict__:
                raise Exception("%s is already in map" % raw_map_key)
            this_section.map.__dict__[raw_map_key] = raw_map_val

        #parse mandatory
        mandatory_lst = this_raw.mandatory.split(",")
        for mandatory in mandatory_lst:
            if not self.__is_key_word(mandatory):
                raise Exception("mandatory:%s is incorrect format" % mandatory)
            if mandatory not in this_section.map.__dict__:
                raise Exception("mandatory:%s is not exist in value_map" % mandatory)
            if mandatory not in this_section.mandatory:
                this_section.mandatory.append(mandatory)
            else:
                raise Exception("mandatory:%s is repeated" % mandatory)

        #parse is_structure
        if this_raw.is_structure not in ["yes","no"]:
            raise Exception("is_structure only have two pattern:yes/no")
        if this_raw.is_structure == "yes":
            this_section.is_struct = True

    def __init__(self,ini_file):
        self.conf = ConfigParser()
        self.__column_range_regex = re.compile("^\s*(\(|\[)\s*(\d+)\s*(,)\s*(\d+)\s*(\)|\])")
        self.conf.read(ini_file)

        self.check_key_lst = [" ","\n","\"","\'"]
        self.sections = self.conf.sections()

        self.must_keys = ["column_range","key_list",
                               "value_map","mandatory","is_structure"]
        for section in self.sections:
            self.__assign_mandatory_attribute(section, self.must_keys, self.conf)
            self.__parse_section(section)
