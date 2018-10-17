Imports Microsoft.VisualBasic
Imports System.Linq
Imports System.Data.Linq
Imports System.IO
Imports System.Net
Imports System
Imports System.Web
Imports System.Collections.Generic

Public Class MyData

  Inherits System.Web.UI.Page

  Private sTest As String = ""

  Private Shared mHouseRentalContext As HouseRentalDataContext = Nothing
  Public Shared ReadOnly Property HouseRentalContext() As HouseRentalDataContext
    Get
      If mHouseRentalContext Is Nothing Then
        mHouseRentalContext = GetHouseRentalContext()
      End If
      Return mHouseRentalContext
    End Get
  End Property

  Private Shared mHouseRentalContextTracking As HouseRentalDataContext = Nothing
  Public Shared ReadOnly Property HouseRentalContextTracking() As HouseRentalDataContext
    Get
      If mHouseRentalContextTracking Is Nothing Then
        mHouseRentalContextTracking = GetHouseRentalContext(True)
      End If
      Return mHouseRentalContextTracking
    End Get
  End Property
  Public Shared Function GetHouseRentalContext(Optional ByVal bEnableObjectTracking As Boolean = False, Optional sConnString As String = "") As HouseRentalDataContext
    GetHouseRentalContext = Nothing
    Dim oLocalContext As HouseRentalDataContext
    If sConnString <> "" Then
      oLocalContext = New HouseRentalDataContext(sConnString)
    Else
      oLocalContext = New HouseRentalDataContext()
    End If
    Try
      oLocalContext.CommandTimeout = 120
      oLocalContext.ObjectTrackingEnabled = bEnableObjectTracking
      GetHouseRentalContext = oLocalContext
    Catch ex As Exception
      oLocalContext.Dispose()
      oLocalContext = Nothing
      Throw New Exception("Error in GetHouseRentalContext", ex)
    End Try

  End Function

  Public Shared Function DisposeHouseRentalContext(ByRef MyDataContext As HouseRentalDataContext, Optional ByVal bSubmitChanges As Boolean = False) As Boolean
    DisposeHouseRentalContext = False
    Try
      If bSubmitChanges Then MyDataContext.SubmitChanges()
      MyDataContext.Connection.Close()
      MyDataContext.Dispose()
      MyDataContext = Nothing
      GC.Collect()
      GC.WaitForPendingFinalizers()
      DisposeHouseRentalContext = True
    Catch ex As Exception

    End Try
  End Function

  ' **** IMPORTANT: This function is duplicated in the MyData class in the reserve and main websites, and the CStayShared DLL component.  Any changes must be copied to all three locations
  Public Shared Function SetPropertyName(ByVal sPropertyName As String, iPropertyCategory As Integer, Optional iSecondCategory As Integer = 0) As String

    SetPropertyName = ""

    Select Case iPropertyCategory
      Case 1  ' Apartment
        SetPropertyName = "apartment-rentals"
      Case 3, 13 ' Lakefront
        SetPropertyName = "house-rentals-on-lake"
      Case 4 ' Riverfront
        SetPropertyName = "house-rentals-on-lake"
      Case 5 ' Rooms and suites
        SetPropertyName = "rooms-suites"
      Case 6, 7, 8, 9, 11 ' Homes
        SetPropertyName = "house-rentals"
      Case 10 ' Oneonta
        SetPropertyName = "oneonta-ny-lodging"
      Case 15 ' Group
        SetPropertyName = "group-lodging"
      Case Else
        SetPropertyName = "lodging-rentals"
    End Select

    ' BJR - changed to not have so much variation in URLs for same content

    If SetPropertyName = "" Then SetPropertyName = "rentals"
    '    If SetPropertyName = "" Then SetPropertyName = sCity & sDreamsPark & sCategory & "-Rental" & sHomeType
    ' If SetPropertyName = "" Then SetPropertyName = sCategory & "/" & sCity & sDreamsPark & "-Rental" & sHomeType
    ' If make change here, need to change GetPropertyInfoByName
    SetPropertyName = SetPropertyName & "/" & sPropertyName.Replace("-", "").Replace(" ", "-").Replace("'", "").Replace("#", "Number").Replace(".", "").Replace("&", "and")

  End Function


  Protected Overrides Sub Finalize()
    If mHouseRentalContext IsNot Nothing Then DisposeHouseRentalContext(mHouseRentalContext)
    If mHouseRentalContextTracking IsNot Nothing Then DisposeHouseRentalContext(mHouseRentalContextTracking)

    MyBase.Finalize()
  End Sub




End Class
