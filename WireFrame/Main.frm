VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##########################################################
'#  Author: Mr KACI Lounes                                #
'#  A 3D WireFrame Engine in *Pure* VB Code!              #
'#  Compile for more speed                                #
'#  Mail me at KLKEANO@HOTMAIL.COM                        #
'#         Copyright © 2005 - KACI Lounes                 #
'##########################################################

Option Explicit            'Stop undeclared variables

Const FOV% = 400           'Field of view, or the distance between
                           'the eye and the projection plane

Const CFar = 0             'ClipFar & ClipNear Zs values
Const CNear = 100

Dim Scene() As Mesh        'Array contains the data structures for all objects

Dim Cam As Vector3D        'Camera vector
Dim ViewMatrix As Matrix   'View matrix
Dim ViewMode As Boolean    'View mode, choose between LookAt(True), or Pitch/Yaw(False)

Dim Normal As Vector3D     'This vector is the normal for each face, instead
                           'of putting load more memory uselessly

Dim LookAtObj%             'Which object to look at ? (1, 2, 3)
Dim XAng!, YAng!           'Anlges rotations (Pitch/Yaw mode)
Dim Pitch!, Yaw!           '          ,,        ,,
Dim I%, J%                 'Loops counters
Dim TeapAng%               'Teapot ZAngle rotation

                           'Booleans for processing keyboard input:

Dim KeyESC As Boolean      'Escape Key     : Exit program
Dim KeySPC As Boolean      'Space Key      : Change view mode
Dim KeyHom As Boolean      'Home Key       : Change X Angle (+)
Dim KeyEnd As Boolean      'End Key        : Change X Angle (-)
Dim KeyLft As Boolean      'Left Key       : Change Y Angle (+)
Dim KeyRgt As Boolean      'Right Key      : Change Y Angle (-)
Dim KeyTop As Boolean      'Up Key         : Move front (Z+)
Dim KeyBot As Boolean      'Down Key       : Move back (Z-)
Dim KeyPad1 As Boolean     'NumPad 1 Key   : Look at the Sphere (LookAt mode only)
Dim KeyPad2 As Boolean     'NumPad 2 Key   : Look at the Teapot (LookAt mode only)
Dim KeyPad3 As Boolean     'NumPad 3 Key   : Look at the Torus (LookAt mode only)
Sub Process()

 'Calculate the view matrix:
 '==========================
 '
 If ViewMode = True Then  'LookAt mode at LookAtObj%:
  Select Case LookAtObj
   Case 1: ViewMatrix = MatrixView(Cam, VectorInput3(30, -10, 0), VectorInput3(0, 1, 0))   'At Sphere
   Case 2: ViewMatrix = MatrixView(Cam, VectorInput3(0, 0, 30), VectorInput3(0, 1, 0))     'At Teapot
   Case 3: ViewMatrix = MatrixView(Cam, VectorInput3(-20, -5, -10), VectorInput3(0, 1, 0)) 'At Torus
  End Select
 Else 'Pitch & Yaw mode:
  Pitch = (Pi * 2) - (XAng * Deg)
  Yaw = (Pi * 2) - (YAng * Deg)
  ViewMatrix = MatrixRotate(0, Pitch)
  ViewMatrix = MatrixMultiply(ViewMatrix, MatrixRotate(1, Yaw))
  ViewMatrix = MatrixMultiply(ViewMatrix, MatrixTranslate(VectorInput3(-Cam.X, -Cam.Y, -Cam.Z)))
 End If

 '##################################################

 'Prepare teapot matrix for ZRotation (Roll):
 '==========================================
 '
 Scene(2).IDMatrix = MatrixWorld(VectorInput3(0, -1, 30), VectorInput3(0.5, 0.5, 0.5), Deg * 90, TeapAng * Deg, 0)

 '##################################################

 'Transform & Project :
 '====================
 '
 For J = LBound(Scene) To UBound(Scene)
  For I = LBound(Scene(J).Vertices) To UBound(Scene(J).Vertices)
   'World transformation:
   Scene(J).TmpVerts(I) = MatrixMultiplyVector3(Scene(J).Vertices(I), Scene(J).IDMatrix)
   'View transformation:
   Scene(J).TmpVerts(I) = MatrixMultiplyVector3(Scene(J).TmpVerts(I), ViewMatrix)

   'Projection (Persective Distortion): (you can change the FOV)
   'Ignore the division by zero, by remplacing it by 0.0001:
   If Scene(J).TmpVerts(I).Z = 0 Then Scene(J).TmpVerts(I).Z = 0.0001
   'Apply the persective distortion ((XY/Z) * FOV),
   'For an orthographic projection, We simply skip the next two lines:
   Scene(J).TmpVerts(I).X = (Scene(J).TmpVerts(I).X / Scene(J).TmpVerts(I).Z) * FOV
   Scene(J).TmpVerts(I).Y = (Scene(J).TmpVerts(I).Y / Scene(J).TmpVerts(I).Z) * FOV
  Next I
 Next J

 '##################################################

 'Hidden faces removal:
 '=====================
 '
 ' 1- Check the visiblity of faces by the normal
 ' 2- Check if the triangle is between CFar & CNear
 '
 For J = LBound(Scene) To UBound(Scene)
  For I = LBound(Scene(J).Faces) To UBound(Scene(J).Faces)
   'Get the face normal:
   Normal = VectorGetNormal(Scene(J).TmpVerts(Scene(J).Faces(I).A), _
                            Scene(J).TmpVerts(Scene(J).Faces(I).B), _
                            Scene(J).TmpVerts(Scene(J).Faces(I).C))

   If Normal.Z > 0 Then
    If Scene(J).TmpVerts(Scene(J).Faces(I).A).Z > CFar And _
       Scene(J).TmpVerts(Scene(J).Faces(I).B).Z > CFar And _
       Scene(J).TmpVerts(Scene(J).Faces(I).C).Z > CFar Then

     If Scene(J).TmpVerts(Scene(J).Faces(I).A).Z < CNear And _
        Scene(J).TmpVerts(Scene(J).Faces(I).B).Z < CNear And _
        Scene(J).TmpVerts(Scene(J).Faces(I).C).Z < CNear Then
      Scene(J).Faces(I).Visible = True
     Else
      Scene(J).Faces(I).Visible = False
     End If

    Else
     Scene(J).Faces(I).Visible = False
    End If

   Else
    Scene(J).Faces(I).Visible = False
   End If

  Next I
 Next J

 TeapAng = TeapAng + 1

End Sub
Sub ClipRay(X1!, Y1!, Z1!, X2!, Y2!, Z2!, OX1!, OY1!, OX2!, OY2!, DrawIt As Byte)

 'Note: this rountine use only the 2D clipping,
 '      in an others next versions, I would to apply
 '      the 3D clipping with Flat-Shading and Gouraud-Shading
 '      (You can also enable the textures !!!), also (!),
 '      You can use the SpotLight filter of my previous
 '      version of my projects! (type this as criteria: KACI Lounes)

 Dim XO1!, YO1!, XO2!, YO2!

 XO1 = X1: YO1 = Y1: XO2 = X2: YO2 = Y2 'Set defaults output values

 '2D Clipping, trivial cases:

 If Accept(-320, -240, 320, 240, XO1, YO1, XO2, YO2) = False Then
  'The line is completly inside the screen, so draw it normaly:
  OX1 = XO1: OY1 = YO1: OX2 = XO2: OY2 = YO2: DrawIt = 1
 Else
  If Reject(-320, -240, 320, 240, XO1, YO1, XO2, YO2) = False Then
   'The line is partialy inside the screen, so clip and draw it:
   ClipLine -320, -240, 320, 240, XO1, YO1, XO2, YO2, OX1, OY1, OX2, OY2
   DrawIt = 1
  End If
  'else: the line is completly outside the screen, so skip it !
 End If

End Sub
Sub GetKeys()

 'Process keyboard entry:
 If KeyESC = True Then Unload Me: End
 If KeySPC = True Then ViewMode = Not ViewMode: KeySPC = False

 If KeyTop = True Then Cam.Z = Cam.Z + 1: KeyTop = False
 If KeyBot = True Then Cam.Z = Cam.Z - 1: KeyBot = False

 If KeyHom = True Then XAng = XAng + 1: KeyHom = False
 If KeyEnd = True Then XAng = XAng - 1: KeyEnd = False
 If KeyRgt = True Then YAng = YAng + 1: KeyRgt = False
 If KeyLft = True Then YAng = YAng - 1: KeyLft = False

 If KeyPad1 = True Then LookAtObj = 1: KeyPad1 = False
 If KeyPad2 = True Then LookAtObj = 2: KeyPad2 = False
 If KeyPad3 = True Then LookAtObj = 3: KeyPad3 = False

End Sub
Sub LoadScene()

 'Redim scene for 4 meshs (Grid, Sphere, Teapot & Torus):
 ReDim Scene(3)

 Dim File3DName$, Buff&  'Buff& is a 'Long' data type, so coded into 4 bytes (32 Bits)
                         'then note that the Get function recieve 4 bytes for each
                         ' 'Long' (or Single) data type readed.

 'Load the models:
 For J = 0 To UBound(Scene)

  Select Case J
   Case 0: File3DName = App.Path & "\Primatives\Grid.klf"
   Case 1: File3DName = App.Path & "\Primatives\Sphere.klf"
   Case 2: File3DName = App.Path & "\Primatives\Teapot.klf"
   Case 3: File3DName = App.Path & "\Primatives\Torus.klf"
  End Select

  '##################################################

  Open File3DName For Binary As 1  'Open the file:
   Get #1, , Buff                  'Number of vertices
   ReDim Scene(J).Vertices(Buff)
   ReDim Scene(J).TmpVerts(Buff)
   Get #1, , Buff                  'Number of faces
   ReDim Scene(J).Faces(Buff)
   For I = LBound(Scene(J).Vertices) To UBound(Scene(J).Vertices) 'Read vertices
    Get #1, , Scene(J).Vertices(I).X
    Get #1, , Scene(J).Vertices(I).Y
    Get #1, , Scene(J).Vertices(I).Z
   Next I
   For I = LBound(Scene(J).Faces) To UBound(Scene(J).Faces)       'Read faces
    Get #1, , Scene(J).Faces(I).A
    Get #1, , Scene(J).Faces(I).B
    Get #1, , Scene(J).Faces(I).C
   Next I
  Close 1                          'Close the file

  '##################################################

  'Set the world matrix for each object in scene:
  Select Case J
   'The grid should be big that others objects (as a floor):
   Case 0: Scene(0).IDMatrix = MatrixWorld(VectorInput3(0, -1, 0), VectorInput3(-1, 1, 1), 0, 0, 0)
   Case 1: Scene(1).IDMatrix = MatrixWorld(VectorInput3(30, -10, 0), VectorInput3(0.3, 0.3, 0.3), 0, 0, 0)
   Case 2: Scene(2).IDMatrix = MatrixWorld(VectorInput3(0, -1, 30), VectorInput3(0.5, 0.5, 0.5), Deg * 90, 0, 0)
   Case 3: Scene(3).IDMatrix = MatrixWorld(VectorInput3(-20, -5, -10), VectorInput3(0.3, 0.3, 0.3), 0, 0, 0)
  End Select

 Next J

 '##################################################

 'Setup the camera:
 Cam = VectorInput3(0.5, -5.5, -30.5): YAng = 15
 ViewMatrix = MatrixIdentity
 ViewMode = False 'Default value: Pitch/Yaw mode
 LookAtObj = 1    'Default value: At the Sphere

End Sub
Sub Render()

 Dim X1!, Y1!, Z1!, X2!, Y2!, Z2!  'Ray to be clipped
 Dim OX1!, OY1!, OX2!, OY2!        'Out screen coordinates
 Dim DIt As Byte

 Line (20, 20)-(660, 500), RGB(0, 50, 50), B 'Draw Clip Boundary (B or BF)

 For J = LBound(Scene) To UBound(Scene)
  Select Case J
   Case 0: ForeColor = vbWhite
   Case 1: ForeColor = vbRed
   Case 2: ForeColor = vbYellow
   Case 3: ForeColor = vbGreen
  End Select
  For I = LBound(Scene(J).Faces) To UBound(Scene(J).Faces)
   If Scene(J).Faces(I).Visible = True Then

    'Render line-by-line:
    '###############################################
    'AB Ray:
    X1 = Scene(J).TmpVerts(Scene(J).Faces(I).A).X
    Y1 = Scene(J).TmpVerts(Scene(J).Faces(I).A).Y
    Z1 = Scene(J).TmpVerts(Scene(J).Faces(I).A).Z
    X2 = Scene(J).TmpVerts(Scene(J).Faces(I).B).X
    Y2 = Scene(J).TmpVerts(Scene(J).Faces(I).B).Y
    Z2 = Scene(J).TmpVerts(Scene(J).Faces(I).B).Z

    ClipRay X1, Y1, Z1, X2, Y2, Z2, OX1, OY1, OX2, OY2, DIt
    If DIt = 1 Then Line (340 + OX1, 260 + OY1)-(340 + OX2, 260 + OY2), ForeColor

    '###############################################
    'BC Ray:
    X1 = Scene(J).TmpVerts(Scene(J).Faces(I).B).X
    Y1 = Scene(J).TmpVerts(Scene(J).Faces(I).B).Y
    Z1 = Scene(J).TmpVerts(Scene(J).Faces(I).B).Z
    X2 = Scene(J).TmpVerts(Scene(J).Faces(I).C).X
    Y2 = Scene(J).TmpVerts(Scene(J).Faces(I).C).Y
    Z2 = Scene(J).TmpVerts(Scene(J).Faces(I).C).Z

    ClipRay X1, Y1, Z1, X2, Y2, Z2, OX1, OY1, OX2, OY2, DIt
    If DIt = 1 Then Line (340 + OX1, 260 + OY1)-(340 + OX2, 260 + OY2), ForeColor

    '###############################################
    'CA Ray:
    X1 = Scene(J).TmpVerts(Scene(J).Faces(I).C).X
    Y1 = Scene(J).TmpVerts(Scene(J).Faces(I).C).Y
    Z1 = Scene(J).TmpVerts(Scene(J).Faces(I).C).Z
    X2 = Scene(J).TmpVerts(Scene(J).Faces(I).A).X
    Y2 = Scene(J).TmpVerts(Scene(J).Faces(I).A).Y
    Z2 = Scene(J).TmpVerts(Scene(J).Faces(I).A).Z

    ClipRay X1, Y1, Z1, X2, Y2, Z2, OX1, OY1, OX2, OY2, DIt
    If DIt = 1 Then Line (340 + OX1, 260 + OY1)-(340 + OX2, 260 + OY2), ForeColor

   End If
  Next I
 Next J

End Sub
Private Sub Form_Activate()

 '#########
 'Main loop
 '#########

 Do
  Cls      'Note: Turn the AutoRedraw off, and have fun !

  GetKeys
  Process
  Render

  DoEvents
 Loop

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 'Generate the keyboard:

 If KeyCode = vbKeyEscape Then KeyESC = True
 If KeyCode = vbKeySpace Then KeySPC = True
 If KeyCode = vbKeyLeft Then KeyLft = True
 If KeyCode = vbKeyRight Then KeyRgt = True
 If KeyCode = vbKeyUp Then KeyTop = True
 If KeyCode = vbKeyDown Then KeyBot = True
 If KeyCode = vbKeyHome Then KeyHom = True
 If KeyCode = vbKeyEnd Then KeyEnd = True
 If KeyCode = vbKeyNumpad1 Then KeyPad1 = True
 If KeyCode = vbKeyNumpad2 Then KeyPad2 = True
 If KeyCode = vbKeyNumpad3 Then KeyPad3 = True

End Sub
Private Sub Form_Load()

 'Redim our window as (640x480), the 680x520
 'resolution is conceived for showing the clipping process.
 Move 0, 0, (680 * 15), (520 * 15)
 ScaleMode = vbPixels

 LoadScene

End Sub
