Option Explicit On
Module GlobalConstants

    'Public m_TrEdit As transitEdit

    'public feature classes
    Public g_Schema As String

    Public Const g_RefEdge = "TransRefEdges"
    Public Const g_RefJunct = "TransRefJunctions"
    Public Const g_TransitLines = "TransitLines"
    Public Const g_TransitPoints = "TransitPoints"
    Public Const g_ModeAttributes = "modeAttributes"
    Public Const g_ModeTolls = "ModeTolls"
    Public Const g_EdgeFac = "tblEdgeFacility"
    Public Const g_TurnMovements = "TurnMovements"
    Public Const g_ProjectRoutes = "ProjectRoutes"
    Public Const g_NamedRoutes = "NamedRoutes"
    Public Const g_BikeAttributes = "BikeAttributes"

    'public fields
    Public Const g_fldJctID = "PSRCJunctID"
    Public Const g_fldEdgeID = "PSRCEDGEID"
    Public Const g_fldINode = "INODE"
    Public Const g_fldJNode = "JNODE"

    Public Const g_fldFromEdge = "FrEdgeID"
    Public Const g_fldToEdge = "ToEdgeID"

    Public Const g_fldTrLineID = "LineID"
    Public Const g_fldTrLineNo = "TransLineNo"
    Public Const g_fldTimePeriod = "TimePeriod"
    Public Const g_fldMode = "Mode"
    Public Const g_fldVehicleType = "VehicleType"
    Public Const g_fldHeadway = "Headway"
    Public Const g_fldSpeed = "Speed"
    Public Const g_fldDescription = "Description"
    Public Const g_fldGeography = "Geography"
    Public Const g_fldOperator = "Operator"
    Public Const g_fldOneway = "Oneway"
    Public Const g_fldInSvcDate = "InServiceDate"
    Public Const g_fldOutSvcDate = "OutServiceDate"
    Public Const g_fldSeats = "Seats"
    Public Const g_fldCapacityVehicles = "CapacityVehicles"
    Public Const g_fldPrjID = "ProjectID"
    Public Const g_fldPrjDB = "ProjectDB"
    Public Const g_fldPath = "Path"

    Public Const g_fldPointOrder = "PointOrder"
    Public Const g_fldTimeFuncId = "TimeFuncID"
    Public Const g_fldDwtStop = "DwtStop"
    Public Const g_fldLayover = "Layover"
    Public Const g_fldUser1 = "User1"
    Public Const g_fldUser2 = "User2"
    Public Const g_fldUser3 = "User3"
    Public Const g_fldUseGPOnly = "UseGPOnly"
    Public Const g_fldIsTimePoint = "IsTimePoint"

    Public Const g_fldFromMpt = "MStart"
    Public Const g_fldToMpt = "MStop"
    Public Const g_fldMpt = "M"
    Public Const g_fldLineEvent = "Line"

    'Transit edits
    Public g_TrLineID As String 'it gets value from the transit line dropdown list
    Public g_TrLineIDNew As String
End Module
