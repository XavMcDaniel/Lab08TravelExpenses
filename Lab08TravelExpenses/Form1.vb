Public Class Form1


    Const decMEAL_REIMBURSEMENT As Decimal = 37D
    Const decPARKING_REIMBURSEMENT As Decimal = 10D
    Const decTAXI_REIMBURSEMENT As Decimal = 20D
    Const decLODGING_REIMBURSEMENT As Decimal = 95D
    Const decMILES_REIMBURSEMENT As Decimal = 0.27D


    Function CalcLodging() As Decimal
        Dim decLodgingReimbursement As Decimal
        decLodgingReimbursement = (CDec(txtDays.Text) * decLODGING_REIMBURSEMENT)
        Return decLodgingReimbursement
    End Function

    Function CalcMeals() As Decimal
        Dim decMealReimbursement As Decimal
        decMealReimbursement = (CDec(txtDays.Text) * decMEAL_REIMBURSEMENT)
        Return decMealReimbursement
    End Function

    Function CalcMileage() As Decimal
        Dim decMileageReimbursement As Decimal
        decMileageReimbursement = (CDec(txtMiles.Text) * decMILES_REIMBURSEMENT)
        Return decMileageReimbursement
    End Function

    Function CalcParkingFees() As Decimal
        Dim decParkingReimbursement As Decimal
        decParkingReimbursement = CDec(txtDays.Text) * decPARKING_REIMBURSEMENT
        Return decParkingReimbursement
    End Function

    Function CalcTaxiFees() As Decimal
        Dim decTaxiReimbursement As Decimal
        decTaxiReimbursement = decTAXI_REIMBURSEMENT * CDec(txtDays.Text)
        Return decTaxiReimbursement
    End Function

    Function CalcTotalReimbursement() As Decimal
        Dim decTotalReimbursement As Decimal
        decTotalReimbursement = CalcLodging() + CalcTaxiFees() +
        CalcMeals() + CalcParkingFees()
        Return decTotalReimbursement
    End Function

    Function CalcUnallowed() As Decimal
        Dim decUnallowed As Decimal
        decUnallowed = (CDec(txtPark.Text) - CalcParkingFees()) +
                        (CDec(txtTaxi.Text) - CalcTaxiFees()) + (CDec(txtLodge.Text) - CalcLodging()) +
                        (CDec(txtMeals.Text) - CalcMeals())
        Return decUnallowed
    End Function

    Function CalcSaved() As Decimal
        Dim decSaved As Decimal
        decSaved = (CDec(txtDays.Text) * decMEAL_REIMBURSEMENT - CDec(txtMeals.Text)) +
                    (CDec(txtDays.Text) * decPARKING_REIMBURSEMENT - CDec(txtPark.Text)) +
                    (CDec(txtDays.Text) * decTAXI_REIMBURSEMENT - CDec(txtTaxi.Text)) +
                    (CDec(txtDays.Text) * decLODGING_REIMBURSEMENT - CDec(txtLodge.Text))
        Return decSaved
    End Function

    Function CalcTotal() As Decimal
        Dim decTotal As Decimal
        decTotal = CDec(txtAirfare.Text) + CDec(txtSeminar.Text) + CDec(txtMeals.Text) +
            (CDec(txtMiles.Text) * decMILES_REIMBURSEMENT) + CDec(txtCarRental.Text) +
            CDec(txtLodge.Text) + CDec(txtPark.Text) + CDec(txtTaxi.Text)
        Return decTotal
    End Function

    Function InputNumeric() As Boolean
        Dim blnNumeric As Boolean
        If IsNumeric(txtAirfare.Text) And
            IsNumeric(txtCarRental.Text) And
            IsNumeric(txtDays.Text) And
            IsNumeric(txtLodge.Text) And
            IsNumeric(txtMeals.Text) And
            IsNumeric(txtMiles.Text) And
            IsNumeric(txtPark.Text) And
            IsNumeric(txtSeminar.Text) And
            IsNumeric(txtTaxi.Text) Then
            blnNumeric = True
        End If
        Return blnNumeric
    End Function

    Function InputPositive() As Boolean
        Dim blnPositive As Boolean
        If CDbl(txtAirfare.Text) >= 0 And
        CDbl(txtCarRental.Text) >= 0 And
        CDbl(txtDays.Text) >= 0 And
        CDbl(txtLodge.Text) >= 0 And
        CDbl(txtMeals.Text) >= 0 And
        CDbl(txtMiles.Text) >= 0 And
        CDbl(txtPark.Text) >= 0 And
        CDbl(txtSeminar.Text) >= 0 And
        CDbl(txtCarRental.Text) >= 0 Then
            blnPositive = True
        End If
        Return blnPositive
    End Function

    Sub InputEmpty()
        If txtAirfare.Text = String.Empty Then
            txtAirfare.Text = "0"
        End If
        If txtCarRental.Text = String.Empty Then
            txtCarRental.Text = "0"
        End If
        If txtDays.Text = String.Empty Then
            txtDays.Text = "0"
        End If
        If txtLodge.Text = String.Empty Then
            txtLodge.Text = "0"
        End If
        If txtMeals.Text = String.Empty Then
            txtMeals.Text = "0"
        End If
        If txtMiles.Text = String.Empty Then
            txtMiles.Text = "0"
        End If
        If txtPark.Text = String.Empty Then
            txtPark.Text = "0"
        End If
        If txtSeminar.Text = String.Empty Then
            txtSeminar.Text = "0"
        End If
        If txtTaxi.Text = String.Empty Then
            txtTaxi.Text = "0"
        End If
        If lblExceedings.Text = String.Empty Then
            lblExceedings.Text = "0"
        End If
        If lblSavings.Text = String.Empty Then
            lblSavings.Text = "0"
        End If
        If lblExpenses.Text = String.Empty Then
            lblExpenses.Text = "0"
        End If
        If lblReimbursements.Text = String.Empty Then
            lblReimbursements.Text = "0"
        End If
    End Sub

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click

        InputEmpty()

        If InputNumeric() Then
            If InputPositive() Then

                If CalcTotalReimbursement() > 0 Then
                    lblReimbursements.Text = CStr(CalcTotalReimbursement().ToString("c"))
                Else
                    lblReimbursements.Text = String.Empty
                End If
                If CalcTotal() > 0 Then
                    lblExpenses.Text = CStr(CalcTotal().ToString("c"))
                Else
                    lblExpenses.Text = String.Empty
                End If
                If CalcUnallowed() > 0 Then
                    lblExceedings.Text = CStr(CalcUnallowed().ToString("c"))
                Else
                    lblExceedings.Text = String.Empty
                End If
                If CalcSaved() > 0 Then
                    lblSavings.Text = CStr(CalcSaved().ToString("c"))
                Else
                    lblSavings.Text = String.Empty
                End If
            Else
                MessageBox.Show("Enter a Positive Valid Amount")
            End If
        Else
            MessageBox.Show("Enter a Numeric Amount")
        End If
    End Sub

    Private Sub btnExit_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtAirfare.Text = String.Empty
        txtCarRental.Text = String.Empty
        txtDays.Text = String.Empty
        txtLodge.Text = String.Empty
        txtMeals.Text = String.Empty
        txtMiles.Text = String.Empty
        txtPark.Text = String.Empty
        txtSeminar.Text = String.Empty
        txtTaxi.Text = String.Empty
    End Sub
End Class