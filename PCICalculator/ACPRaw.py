from RawData import RawData
class ACPRaw(RawData):
    """Class used for the rating data from the Raw Data csv"""
    AlligatorL=""
    AlligatorM=""
    AlligatorH=""
    BlockL=""
    BlockM=""
    BlockH=""
    DistortionL=""
    DistortionM=""
    DistortionH=""
    LongTransL=""
    LongTransM=""
    LongTransH=""
    PatchL=""
    PatchM=""
    PatchH=""
    RavelingL=""
    RavelingM=""
    RavelingH=""
    RuttingDepressionL=""
    RuttingDepressionM=""
    RuttingDepressionH=""
    WeatheringL=""
    WeatheringM=""
    WeatheringH=""
    CalcPCI=""

    #Constructor
    def __init__(self,StIDSecID,SampleNum,Rater,StreetName,BegLocation,EndLocation,SampleLength,SampleWidth,Date,SampleNotes,Photos,QA,Special,SampleArea,
                 AllL,AllM,AllH,BlokL,BlokM,BlokH,DisL,DisM,DisH,LTCL,LTCM,LTCH,PatL,PatM,PatH,RavL,RavM,RavH,RutL,RutM,RutH,WeatL,WeatM,WeatH):
        
        #Parent Class Constructor
        super().__init__(StIDSecID,SampleNum,Rater,StreetName,BegLocation,EndLocation,SampleLength,SampleWidth,Date,SampleNotes,Photos,QA,Special,SampleArea)

        #Alligator Ratings
        self.AlligatorL=AllL
        self.AlligatorM=AllM
        self.AlligatorH=AllH

        #Block Ratings
        self.BlockL=BlokL
        self.BlockM=BlokM
        self.BlockH=BlokH

        #Distortion Ratings
        self.DistortionL=DisL
        self.DistortionM=DisM
        self.DistortionH=DisH

        #Long Trans Ratings
        self.LongTransL=LTCL
        self.LongTransM=LTCM
        self.LongTransH=LTCH

        #Patch Ratings
        self.PatchL=PatL
        self.PatchM=PatM
        self.PatchH=PatH

        #Raveling Ratings
        self.RavelingL=RavL
        self.RavelingM=RavM
        self.RavelingH=RavH

        #Rutting Ratings
        self.RuttingDepressionL=RutL
        self.RuttingDepressionM=RutM
        self.RuttingDepressionH=RutH

        #Weathering Ratings
        self.WeatheringL=WeatL
        self.WeatheringM=WeatM
        self.WeatheringH=WeatH