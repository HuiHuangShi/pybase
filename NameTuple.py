class NameTuple:
    def __init__(self,fmt):
        _type = type(fmt)
        if not (_type == list or _type == tuple):
             raise Exception("fmt is not match:%s" % _type)

        self.colMap = {}
        for idx in xrange(0,len(fmt)):
            self.colMap[idx] = fmt[idx]

    def parse(self,data):
        _type = type(data)
        if not (_type == list or _type == tuple):
            raise Exception("data is not match:%s" % _type)

        if len(data) != len(self.colMap):
            raise Exception("data len dismatch")

        ret = BaseTuple()
        for idx in xrange(0,len(data)):
            ret.__dict__[self.colMap[idx]] = data[idx]

        return ret

    def transfer(self,base,isList=False):
        ret = []
        for idx in xrange(0,len(self.colMap)):
            if self.colMap[idx] not in base.__dict__:
                raise Exception("base do not have member:%s" % self.colMap[idx])
            ret.append(base.__dict__[self.colMap[idx]])
        if isList:
            return ret
        else:
            return tuple(ret)
