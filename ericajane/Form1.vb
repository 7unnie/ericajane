Public Class Form1


    Dim QP1, QP2, QP3, ATTP, RP1, RP2, ACTP1, ACTP2, ACTP3, EXAMP As Double
    Dim TQP, TATTP, TRP, TACTP, TEXAMP, PRELIM As Double
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        QP1 = Val(TextBox1.Text)
        QP2 = Val(TextBox2.Text)
        QP3 = Val(TextBox3.Text)
        ATTP = Val(TextBox4.Text)
        RP1 = Val(TextBox5.Text)
        RP2 = Val(TextBox6.Text)
        ACTP1 = Val(TextBox7.Text)
        ACTP2 = Val(TextBox8.Text)
        ACTP3 = Val(TextBox9.Text)
        EXAMP = Val(TextBox10.Text)

        TQP = ((QP1 + QP2 + QP3) / 3) * 0.25
        TATTP = ATTP * 0.1
        TRP = ((RP1 + RP2) / 2) * 0.1
        TACTP = ((ACTP1 + ACTP2 + ACTP3) / 3) * 0.25
        TEXAMP = EXAMP * 0.3

        PRELIM = (TQP + TATTP + TRP + TACTP + TEXAMP) * 0.3

        TextBox31.Text = PRELIM

    End Sub

    ' midterm

    Dim QM1, QM2, QM3, ATTM, RM1, RM2, ACTM1, ACTM2, ACTM3, EXAMM As Double
    Dim TQM, TATTM, TRM, TACTM, TEXAMM, MIDTERM As Double
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        QM1 = Val(TextBox11.Text)
        QM2 = Val(TextBox12.Text)
        QM3 = Val(TextBox13.Text)
        ATTM = Val(TextBox14.Text)
        RM1 = Val(TextBox15.Text)
        RM2 = Val(TextBox16.Text)
        ACTM1 = Val(TextBox17.Text)
        ACTM2 = Val(TextBox18.Text)
        ACTM3 = Val(TextBox19.Text)
        EXAMM = Val(TextBox20.Text)

        TQM = ((QM1 + QM2 + QM3) / 3) * 0.25
        TATTM = ATTM * 0.1
        TRM = ((RM1 + RM2) / 2) * 0.1
        TACTM = ((ACTM1 + ACTM2 + ACTM3) / 3) * 0.25
        TEXAMM = EXAMM * 0.3

        MIDTERM = (TQM + TATTM + TRM + TACTM + TEXAMM) * 0.3

        TextBox32.Text = MIDTERM

    End Sub

    'FINALS
    Dim QF1, QF2, QF3, ATTF, RF1, RF2, ACTF1, ACTF2, ACTF3, EXAMF As Double
    Dim TQF, TATTF, TRF, TACTF, TEXAMF, FINALS As Double
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        QF1 = Val(TextBox21.Text)
        QF2 = Val(TextBox22.Text)
        QF3 = Val(TextBox23.Text)
        ATTF = Val(TextBox24.Text)
        RF1 = Val(TextBox25.Text)
        RF2 = Val(TextBox26.Text)
        ACTF1 = Val(TextBox27.Text)
        ACTF2 = Val(TextBox28.Text)
        ACTF3 = Val(TextBox29.Text)
        EXAMF = Val(TextBox30.Text)

        TQF = ((QF1 + QF2 + QF3) / 3) * 0.25
        TATTF = ATTF * 0.1
        TRF = ((RF1 + RF2) / 2) * 0.1
        TACTF = ((ACTF1 + ACTF2 + ACTF3) / 3) * 0.25
        TEXAMF = EXAMF * 0.3

        FINALS = (TQF + TATTF + TRF + TACTF + TEXAMF) * 0.4

        TextBox33.Text = FINALS
    End Sub

    Dim TGRADE As Double

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TGRADE = PRELIM + MIDTERM + FINALS
        TextBox34.Text = TGRADE
    End Sub


End Class
