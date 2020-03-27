class RawData(object):
    """Parent Class for data from the RawData csv"""

    StIDSecID =""
    SampleNumber = ""
    Rater = ""
    StreetName = ""
    BegLocation = ""
    EndLocation = ""
    SampleLength = ""
    SampleWidth = ""
    Date = ""
    SampleNotes = ""
    Photos = ""
    QA = ""
    Special = ""
    SampleArea = ""

    #Constructor
    def __init__(self,StidSecid,SampleNum,Rate,StName,Beg,End,SampleL,SampleWid,Dates,SampleN,Photo,qa,Spec,SampleA):

        self.StIDSecID=StidSecid    
        self.SampleNumber=SampleNum
        self.Rater=Rate
        self.StreetName=StName
        self.BegLocation=Beg
        self.EndLocation=End
        self.SampleLength=SampleL
        self.SampleWidth=SampleWid
        self.Date=Dates
        self.SampleNotes=SampleN
        self.Photos=Photo
        self.QA=qa
        self.Special=Spec
        self.SampleArea=SampleA


