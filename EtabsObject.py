import comtypes.client
class EtabsProject():
    def __init__(self):
        try:
            ActiveObject = True
            myEtabsObject = comtypes.client.GetActiveObject('CSI.ETABS.API.ETABSObject')
            print('Açık bir Etabs projesi bulundu...')
        except:
            ActiveObject = False
            print('Yeni bir proje başlatılıyor...')
            myEtabsObject = comtypes.client.CreateObject('ETABSv1.Helper').QueryInterface(comtypes.gen.ETABSv1.cHelper).\
                CreateObjectProgID('CSI.ETABS.API.ETABSObject')
            myEtabsObject.ApplicationStart()

        finally:
            self.SapModel = myEtabsObject.SapModel
            if ActiveObject == False:
                self.SapModel.InitializeNewModel()

        # self.CreateGridSystem()
        # self.matNames()
        # self.getMatProp()
        # self.createNewMaterial()
        self.deleteMaterial()


    def CreateGridSystem(self):
        self.SapModel.SetPresentUnits(6)
        self.SapModel.File.NewGridOnly(4, 3, 3, 4, 3, 3, 4)
        # self.SapModel.File.NewGridOnly(numberStory, typicalStoryHeight, BottomStoryHeight, numberLinesX, numberLinesY, SpacingX, SpacingY)


    def matNames(self):
        matName = self.SapModel.PropMaterial.GetNameList()
        print(matName)

    def getMatProp(self):
        getProp = self.SapModel.PropMaterial.GetOConcrete('C25/30')
        # GetOConcrete(Name, Fc, IsLightWeight, FcsFactor, SSType, SSHysType, StrainAtFc, StrainUltimate, FrictionAngle, DilatationAngle, temp)
        print(getProp)
        getWeight = self.SapModel.PropMaterial.GetWeightAndMass('C25/30')
        print(getWeight)
        getMechProp = self.SapModel.PropMaterial.GetMPIsotropic('C25/30')
        print(getMechProp)

    def createNewMaterial(self):
        self.SapModel.PropMaterial.SetMaterial('C30', 2)
        # SetOConcrete(Name, Fc, IsLightWeight, FcsFactor, SSType, SSHysType, StrainAtFc, StrainUltimate, FrictionAngle, DilatationAngle, temp)
        # Hognestad Method ==> StrainUltimate = 0.0038 ; Etabs StrainUltimate = 0.005
        # StrainAtFc = 2*Fc/E ==> 2 * 30 / 32000 = 0.001875
        self.SapModel.PropMaterial.SetOConcrete('C30', 30000, False, 0, 2, 4, 0.001875, 0.0038, 0, 0, 0)
        self.SapModel.PropMaterial.SetWeightAndMass('C30', 1, 25)
        self.SapModel.PropMaterial.SetMPIsotropic('C30', 32000000, 0.2, 1e-5)

    def deleteMaterial(self):
        self.SapModel.PropMaterial.Delete('C25/30')


EtabsProject()