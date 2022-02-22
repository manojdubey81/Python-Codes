#**********************************************************************************************************************************************************
#   SkillCategory USER DEFINED FUCTION
#**********************************************************************************************************************************************************

def SkillCategory(SkillName):

        SkillNameUpper = SkillName.upper()
        skills = "MAINFRAME, DOT NET, JAVA, DBA, PM, SME, BA, DOC SPL"
        if skills.find(SkillNameUpper) >= 0:
                SkillCategory = SkillNameUpper
        else:
                SkillCategory = "OTHER"

        return SkillCategory

#**********************************************************************************************************************************************************
#   Skillname USER DEFINED FUNCTION
#**********************************************************************************************************************************************************

def skillname(ResourceName):

        ResourceNameUpper = ResourceName.upper()
        if ResourceNameUpper.find('.ARCH') >= 0:
                SkillName = "ARCHITECT"
        elif ResourceNameUpper.find('.BO') >= 0:
                SkillName = "BO"
        elif ResourceNameUpper.find('.DB') >= 0:
                SkillName = "DBA"
        elif ResourceNameUpper.find('.NET') >= 0:
                SkillName = "DOT NET"
        elif ResourceNameUpper.find('.VB') >= 0:
                SkillName = "DOT NET"
        elif ResourceNameUpper.find('.ETL') >= 0:
                SkillName = "ETL"
        elif ResourceNameUpper.find('.JAVA') >= 0:
                SkillName = "JAVA"
        elif ResourceNameUpper.find('.MF') >= 0:
                SkillName = "MAINFRAME"
        elif ResourceNameUpper.find('.PM') >= 0:
                SkillName = "PM"
        elif ResourceNameUpper.find('.SME') >= 0:
                SkillName = "SME"
        elif ResourceNameUpper.find('.SURGE') >= 0:
                SkillName = "SURGE"
        elif ResourceNameUpper.find('.USOFT') >= 0:
                SkillName = "USOFT"
        elif ResourceNameUpper.find('.RAIS') >= 0:
                SkillName = "USOFT"
        elif ResourceNameUpper.find('.ORACLEDBA') >= 0:
                SkillName = "ORACLE DBA"
        elif ResourceNameUpper.find('.MCWEB') >= 0:
                SkillName = "DOT NET"
        elif ResourceNameUpper.find('.CMS64') >= 0:
                SkillName = "DOT NET"
        elif ResourceNameUpper.find('.WTX') >= 0:
                SkillName = "WTX"
        elif ResourceNameUpper.find('.ITX') >= 0:
                SkillName = "WTX"
        elif ResourceNameUpper.find('.BA') >= 0:
                SkillName = "BA"
        elif ResourceNameUpper.find('.DOC SPL') >= 0:
                SkillName = "DOC SPL"
        elif ResourceNameUpper.find('.BUSINESS CONSULTANT') >= 0:
                SkillName = "SME"
        else:
                SkillName = "OTHERS"

        return SkillName



