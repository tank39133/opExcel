class Product(object):
    def __init__(self,id,percentage,contentList,nameList):
        object.__init__(self)
        self.id = id
        self.percentage = percentage
        self.contentList = contentList
        self.nameList = nameList


class Content(object):
    def __init__(self,name_cn,inci,percentage,perInAll):
        object.__init__(self)
        self.name_cn = name_cn
        self.inci = inci
        self.percentage = percentage
        self.perInAll = perInAll