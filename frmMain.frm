VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Power Tetris"
   ClientHeight    =   7815
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3810
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3600
      Top             =   3000
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   7215
      Left            =   100
      ScaleHeight     =   7155
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   100
      Width           =   3615
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   286
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   316
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   285
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   315
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   284
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   314
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   283
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   313
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   282
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   312
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   281
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   311
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   280
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   310
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   279
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   309
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   278
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   308
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   277
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   307
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   276
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   306
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   275
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   305
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   274
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   304
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   273
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   303
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   272
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   302
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   271
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   301
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   270
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   300
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   269
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   299
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   268
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   298
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   267
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   297
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   266
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   296
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   265
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   295
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   264
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   294
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   263
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   293
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   262
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   292
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   261
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   291
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   260
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   290
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   259
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   289
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   258
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   288
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   257
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   287
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   256
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   286
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   255
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   285
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   254
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   284
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   253
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   283
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   252
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   282
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   251
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   281
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   250
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   280
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   249
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   279
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   248
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   278
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   247
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   277
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   246
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   276
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   245
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   275
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   244
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   274
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   243
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   273
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   242
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   272
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   241
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   271
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   240
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   270
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   239
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   269
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   238
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   268
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   237
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   267
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   236
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   266
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   235
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   265
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   234
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   264
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   233
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   263
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   232
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   262
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   231
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   261
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   230
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   260
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   229
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   259
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   228
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   258
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   225
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   257
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   220
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   256
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   219
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   255
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   218
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   254
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   217
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   253
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   216
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   252
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   215
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   251
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   214
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   250
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   213
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   249
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   212
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   248
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   211
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   247
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   210
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   246
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   209
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   245
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   208
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   244
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   207
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   243
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   206
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   242
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   205
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   241
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   204
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   240
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   203
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   239
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   202
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   238
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   201
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   237
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   200
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   236
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   199
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   235
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   198
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   234
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   197
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   233
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   196
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   232
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   195
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   231
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   194
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   230
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   193
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   229
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   192
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   228
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   191
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   227
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   190
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   226
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   189
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   225
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   188
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   224
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   187
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   223
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   186
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   222
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   185
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   221
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   184
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   220
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   183
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   219
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   182
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   218
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   181
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   217
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   180
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   216
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   179
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   215
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   178
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   214
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   177
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   213
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   176
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   212
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   175
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   211
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   174
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   210
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   173
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   209
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   172
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   208
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   171
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   207
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   170
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   206
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   169
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   205
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   168
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   204
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   167
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   203
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   166
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   202
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   165
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   201
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   164
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   200
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   163
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   199
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   162
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   198
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   161
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   197
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   160
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   196
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   159
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   195
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   158
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   194
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   157
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   193
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   156
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   192
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   155
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   191
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   154
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   190
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   153
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   189
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   152
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   188
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   151
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   187
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   150
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   186
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   149
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   185
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   148
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   184
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   147
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   183
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   146
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   182
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   145
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   181
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   144
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   180
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   143
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   179
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   142
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   178
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   141
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   177
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   140
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   176
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   139
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   175
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   138
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   174
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   137
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   173
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   136
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   172
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   135
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   171
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   134
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   170
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   133
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   169
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   132
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   168
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   131
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   167
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   130
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   166
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   129
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   165
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   128
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   164
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   127
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   163
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   126
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   162
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   125
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   161
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   124
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   160
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   123
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   159
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   122
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   158
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   121
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   157
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   120
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   156
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   119
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   155
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   118
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   154
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   117
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   153
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   116
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   152
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   115
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   151
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   114
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   150
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   113
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   149
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   112
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   148
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   111
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   147
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   110
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   146
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   109
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   145
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   108
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   144
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   107
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   143
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   106
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   142
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   105
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   141
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   104
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   140
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   103
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   139
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   102
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   138
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   101
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   137
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   100
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   136
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   99
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   135
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   98
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   134
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   97
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   133
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   96
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   132
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   95
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   131
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   94
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   130
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   93
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   129
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   92
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   128
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   91
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   127
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   90
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   126
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   89
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   125
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   88
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   124
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   87
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   123
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   86
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   122
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   85
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   121
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   84
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   120
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   83
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   119
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   82
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   118
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   81
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   117
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   80
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   116
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   79
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   115
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   78
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   114
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   77
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   113
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   76
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   112
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   75
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   111
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   74
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   110
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   73
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   109
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   72
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   108
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   71
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   107
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   70
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   106
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   69
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   105
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   68
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   104
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   67
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   103
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   66
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   102
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   65
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   101
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   64
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   100
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   63
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   99
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   62
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   98
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   61
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   97
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   60
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   96
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   59
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   95
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   58
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   94
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   57
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   93
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   56
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   92
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   55
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   91
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   54
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   90
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   53
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   89
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   52
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   88
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   51
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   87
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   50
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   86
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   49
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   85
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   48
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   84
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   47
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   83
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   46
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   82
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   45
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   81
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   44
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   80
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   43
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   79
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   42
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   78
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   41
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   77
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   40
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   76
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   39
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   75
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   38
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   74
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   37
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   73
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   36
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   72
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   35
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   71
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   34
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   70
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   33
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   69
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   32
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   68
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   31
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   67
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   30
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   66
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   29
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   65
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   28
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   64
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   27
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   63
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   26
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   62
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   25
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   61
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   24
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   60
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   23
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   59
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   22
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   58
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   21
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   57
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   20
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   56
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   19
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   55
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   18
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   54
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   17
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   53
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   16
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   52
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   15
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   51
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   14
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   50
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   13
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   49
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   12
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   48
         Top             =   6600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   11
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   47
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   46
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   45
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   44
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   43
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   42
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   41
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   1200
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   40
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   39
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   38
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   37
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   36
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   224
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   9
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   223
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   8
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   222
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   7
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   221
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   6
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   227
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   5
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   226
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   4
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   287
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   3
         Top             =   6300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox D3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   1
         Top             =   1500
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox D4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   2
         Top             =   1500
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox A4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   10
         Top             =   300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox A3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox A2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   12
         Top             =   300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox A1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox D1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1500
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox D2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   15
         Top             =   1200
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox C1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   16
         Top             =   300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox C3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   17
         Top             =   300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox C2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   18
         Top             =   300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox C4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   19
         Top             =   300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox B2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   20
         Top             =   1200
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox B4 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   21
         Top             =   900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox B3 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   22
         Top             =   900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox B1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2700
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   23
         Top             =   1200
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox F3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   24
         Top             =   1200
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox F2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   25
         Top             =   900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox F1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2400
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox F4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox E3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox E2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   29
         Top             =   300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox E1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox E4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox G1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox G2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox G3 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   34
         Top             =   300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox G4 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   300
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   35
         Top             =   300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblPause 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "PAUSE..."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   24
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   915
         Left            =   398
         TabIndex        =   318
         Top             =   2700
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3600
      Top             =   1800
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Power Tetris"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   100
      TabIndex        =   317
      Top             =   7400
      Width           =   3615
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu new 
         Caption         =   "&New"
         Index           =   11
         Shortcut        =   {F2}
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
         Index           =   12
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Index           =   2
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************
'in game is 7 different objects represented with letters A,B,...
'each object consists from 4 PictureBox controls
'(for object A --> A1,A2,A3,A4)
'because of different shapes of objects it is neccesary to
'have different functions for rotation and movement
'***************************************************************

Option Explicit

Const NumberOfFields = 287
Const x = 300

Public NumberOfRows As Integer
Public DisplayHighScores As Boolean

Dim Speed   As Integer
Dim Matrica(11, 23) As Boolean
Dim MatricaFull(23) As Boolean
Dim Active_Object As String
Dim Object_A_Position As Integer
Dim Object_B_Position As Integer
Dim Object_C_Position As Integer
Dim Object_D_Position As Integer
Dim Object_E_Position As Integer
Dim Object_F_Position As Integer
Dim Object_G_Position As Integer
Dim Pause As Boolean
Dim GameRunnig As Boolean

Private Sub Rotate_Any_Object()
    Select Case Active_Object
    Case "A"
        Call Rotate_Object_A
    Case "B"
        Call Rotate_Object_B
    Case "C"
        Call Rotate_Object_C
    Case "D"
        Call Rotate_Object_D
    Case "E"
        Call Rotate_Object_E
    Case "F"
        Call Rotate_Object_F
    End Select
End Sub

Private Sub Rotate_Object_A()
On Error GoTo ErrorHandler
    Select Case Object_A_Position
    Case 1
        If (Matrica((A1.Left) / x, (A1.Top + 2 * x) / x) = False And _
        Matrica((A2.Left + x) / x, (A2.Top + x) / x) = False And _
        Matrica((A3.Left + 2 * x) / x, (A3.Top) / x) = False) Then
            A1.Top = A1.Top + 2 * x
            A2.Left = A2.Left + x
            A2.Top = A2.Top + x
            A3.Left = A3.Left + 2 * x
            Object_A_Position = 2
        End If
    Case 2
        If (Matrica((A1.Left + 2 * x) / x, (A1.Top - 2 * x) / x) = False And _
        Matrica((A2.Left + x) / x, (A2.Top - x) / x) = False) Then
            A1.Top = A1.Top - 2 * x
            A1.Left = A1.Left + 2 * x
            A2.Left = A2.Left + x
            A2.Top = A2.Top - x
            Object_A_Position = 3
        End If
    Case 3
        If (Matrica((A1.Left - 2 * x) / x, (A1.Top) / x) = False And _
        Matrica((A2.Left - x) / x, (A2.Top - x) / x) = False And _
        Matrica(A3.Left / x, (A3.Top - 2 * x) / x) = False) Then
            A1.Left = A1.Left - 2 * x
            A2.Left = A2.Left - x
            A2.Top = A2.Top - x
            A3.Top = A3.Top - 2 * x
        Object_A_Position = 4
        End If
    Case 4
        If (Matrica((A2.Left - x) / x, (A2.Top + x) / x) = False And _
        Matrica((A3.Left - 2 * x) / x, (A3.Top + 2 * x) / x) = False) Then
            A2.Left = A2.Left - x
            A2.Top = A2.Top + x
            A3.Top = A3.Top + 2 * x
            A3.Left = A3.Left - 2 * x
            Object_A_Position = 1
        End If
    End Select
ErrorHandler:
End Sub

Private Sub Rotate_Object_B()
On Error GoTo ErrorHandler
    Select Case Object_B_Position
    Case 1
        If (Matrica((B1.Left + 2 * x) / x, (B1.Top + x) / x) = False And _
        Matrica((B4.Left) / x, (B4.Top + x) / x) = False) Then
            B1.Left = B1.Left + 2 * x
            B1.Top = B1.Top + x
            B4.Top = B4.Top + x
            Object_B_Position = 2
        End If
    Case 2
        If (Matrica((B1.Left - 2 * x) / x, (B1.Top - x) / x) = False And _
        Matrica((B4.Left) / x, (B4.Top - x) / x) = False) Then
            B1.Left = B1.Left - 2 * x
            B1.Top = B1.Top - x
            B4.Top = B4.Top - x
            Object_B_Position = 1
        End If
    End Select
ErrorHandler:
End Sub

Private Sub Rotate_Object_C()
On Error GoTo ErrorHandler
    Select Case Object_C_Position
    Case 1
        If (Matrica((C1.Left + x) / x, (C1.Top - x) / x) = False And _
        Matrica((C3.Left - x) / x, (C3.Top + x) / x) = False And _
        Matrica((C4.Left - 2 * x) / x, (C4.Top + 2 * x) / x) = False) Then
            C1.Top = C1.Top - x
            C1.Left = C1.Left + x
            C3.Top = C3.Top + x
            C3.Left = C3.Left - x
            C4.Top = C4.Top + 2 * x
            C4.Left = C4.Left - 2 * x
            Object_C_Position = 2
        End If
    Case 2
        If (Matrica((C1.Left - x) / x, (C1.Top + x) / x) = False And _
        Matrica((C3.Left + x) / x, (C3.Top - x) / x) = False And _
        Matrica((C4.Left + 2 * x) / x, (C4.Top - 2 * x) / x) = False) Then
            C1.Top = C1.Top + x
            C1.Left = C1.Left - x
            C3.Top = C3.Top - x
            C3.Left = C3.Left + x
            C4.Top = C4.Top - 2 * x
            C4.Left = C4.Left + 2 * x
            Object_C_Position = 1
        End If
    End Select
ErrorHandler:
End Sub

Private Sub Rotate_Object_D()
On Error GoTo ErrorHandler
    Select Case Object_D_Position
    Case 1
        If (Matrica((D1.Left + 2 * x) / x, (D1.Top - x) / x) = False And _
        Matrica((D4.Left) / x, (D4.Top - x) / x) = False) Then
            D1.Top = D1.Top - x
            D1.Left = D1.Left + 2 * x
            D4.Top = D4.Top - x
            Object_D_Position = 2
        End If
    Case 2
        If (Matrica((D1.Left - 2 * x) / x, (D1.Top + x) / x) = False And _
        Matrica((D4.Left) / x, (D4.Top + x) / x) = False) Then
            D1.Top = D1.Top + x
            D1.Left = D1.Left - 2 * x
            D4.Top = D4.Top + x
            Object_D_Position = 1
        End If
    End Select
ErrorHandler:
End Sub

Private Sub Rotate_Object_E()
On Error GoTo ErrorHandler
    Select Case Object_E_Position
    Case 1
        If (Matrica((E1.Left - x) / x, (E1.Top + x) / x) = False And _
        Matrica((E3.Left + x) / x, (E3.Top - x) / x) = False And _
        Matrica((E4.Left - 2 * x) / x, (E4.Top) / x) = False) Then
            E1.Top = E1.Top + x
            E1.Left = E1.Left - x
            E3.Top = E3.Top - x
            E3.Left = E3.Left + x
            E4.Left = E4.Left - 2 * x
            Object_E_Position = 2
        End If
    Case 2
        If (Matrica((E1.Left + x) / x, (E1.Top + x) / x) = False And _
        Matrica((E3.Left - x) / x, (E3.Top - x) / x) = False And _
        Matrica((E4.Left) / x, (E4.Top + 2 * x) / x) = False) Then
            E1.Top = E1.Top + x
            E1.Left = E1.Left + x
            E3.Left = E3.Left - x
            E3.Top = E3.Top - x
            E4.Top = E4.Top + 2 * x
            Object_E_Position = 3
        End If
    Case 3
        If (Matrica((E1.Left) / x, (E1.Top - x) / x) = False And _
        Matrica((E2.Left + x) / x, (E2.Top) / x) = False And _
        Matrica((E4.Left + 3 * x) / x, (E4.Top) / x) = False And _
        Matrica((E3.Left + 2 * x) / x, (E3.Top + x) / x) = False) Then
            E1.Top = E1.Top - x
            E2.Left = E2.Left + x
            E3.Left = E3.Left + 2 * x
            E3.Top = E3.Top + x
            E4.Left = E4.Left + 3 * x
            Object_E_Position = 4
        End If
    Case 4
        If (Matrica((E1.Left) / x, (E1.Top - x) / x) = False And _
        Matrica((E2.Left - x) / x, (E2.Top) / x) = False And _
        Matrica((E4.Left - x) / x, (E4.Top - 2 * x) / x) = False And _
        Matrica((E3.Left - 2 * x) / x, (E3.Top + x) / x) = False) Then
            E1.Top = E1.Top - x
            E2.Left = E2.Left - x
            E3.Left = E3.Left - 2 * x
            E3.Top = E3.Top + x
            E4.Left = E4.Left - x
            E4.Top = E4.Top - 2 * x
            Object_E_Position = 1
        End If
    End Select
ErrorHandler:
End Sub

Private Sub Rotate_Object_F()
On Error GoTo ErrorHandler
    Select Case Object_F_Position
    Case 1
        If (Matrica((F1.Left - x) / x, (F1.Top + x) / x) = False And _
        Matrica((F3.Left + x) / x, (F3.Top - x) / x) = False And _
        Matrica((F4.Left) / x, (F4.Top + 2 * x) / x) = False) Then
            F1.Top = F1.Top + x
            F1.Left = F1.Left - x
            F3.Top = F3.Top - x
            F3.Left = F3.Left + x
            F4.Top = F4.Top + 2 * x
            Object_F_Position = 2
        End If
    Case 2
        If (Matrica((F1.Left + x) / x, (F1.Top - x) / x) = False And _
        Matrica((F3.Left - x) / x, (F3.Top + x) / x) = False And _
        Matrica((F4.Left + 2 * x) / x, (F4.Top) / x) = False) Then
            F1.Top = F1.Top - x
            F1.Left = F1.Left + x
            F3.Top = F3.Top + x
            F3.Left = F3.Left - x
            F4.Left = F4.Left + 2 * x
            Object_F_Position = 3
        End If
    Case 3
        If (Matrica((F1.Left - x) / x, (F1.Top + x) / x) = False And _
        Matrica((F3.Left + x) / x, (F3.Top - x) / x) = False And _
        Matrica((F4.Left) / x, (F4.Top - 2 * x) / x) = False) Then
            F1.Top = F1.Top + x
            F1.Left = F1.Left - x
            F3.Top = F3.Top - x
            F3.Left = F3.Left + x
            F4.Top = F4.Top - 2 * x
            Object_F_Position = 4
        End If
    Case 4
        If (Matrica((F1.Left + x) / x, (F1.Top - x) / x) = False And _
        Matrica((F3.Left - x) / x, (F3.Top + x) / x) = False And _
        Matrica((F4.Left - 2 * x) / x, (F4.Top) / x) = False) Then
            F1.Top = F1.Top - x
            F1.Left = F1.Left + x
            F3.Top = F3.Top + x
            F3.Left = F3.Left - x
            F4.Left = F4.Left - 2 * x
            Object_F_Position = 1
        End If
    End Select
ErrorHandler:
End Sub

Private Sub Initialise()
    Dim i As Integer
    Dim j As Integer
        
    Speed = 500
    Timer1.Interval = Speed
    lblStatus.ForeColor = vbGreen
    Timer1.Enabled = False
    Timer2.Enabled = False
    NumberOfRows = 0
   
    A1.Visible = False
    A2.Visible = False
    A3.Visible = False
    A4.Visible = False
    B1.Visible = False
    B2.Visible = False
    B3.Visible = False
    B4.Visible = False
    C1.Visible = False
    C2.Visible = False
    C3.Visible = False
    C4.Visible = False
    D1.Visible = False
    D2.Visible = False
    D3.Visible = False
    D4.Visible = False
    E1.Visible = False
    E2.Visible = False
    E3.Visible = False
    E4.Visible = False
    F1.Visible = False
    F2.Visible = False
    F3.Visible = False
    F4.Visible = False
    G1.Visible = False
    G2.Visible = False
    G3.Visible = False
    G4.Visible = False
    
    Object_A_Position = 1
    Object_B_Position = 1
    Object_C_Position = 1
    Object_D_Position = 1
    Object_E_Position = 1
    Object_F_Position = 1
    Object_G_Position = 1
    
    For i = 0 To 11
        For j = 0 To 23
            Matrica(i, j) = False
        Next j
    Next i
    
    For i = 0 To NumberOfFields
        Pic(i).Visible = False
    Next i
    GameRunnig = True
    Timer1.Enabled = True
End Sub

Private Sub exit_Click(Index As Integer)
    End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If GameRunnig Then
        If Not Pause Then
            If KeyAscii = Asc(5) Then Call Rotate_Any_Object
            If KeyAscii = Asc(4) Then Call Move_Any_Object_Left
            If KeyAscii = Asc(6) Then Call Move_Any_Object_Right
            If KeyAscii = Asc(2) Then Call Move_Any_Object_Down
            If KeyAscii = Asc("+") And Speed > 0 Then Speed = Speed - 10
        End If
        If KeyAscii = Asc("p") Or KeyAscii = Asc("P") Then
            If ((Pause = False) And (GameRunnig = True)) Then
                Timer1.Enabled = False
                lblPause.Visible = True
                Pause = True
            Else
                Timer1.Enabled = True
                lblPause.Visible = False
                Pause = False
            End If
        End If
    End If
End Sub

Private Sub Move_Any_Object_Left()
    Select Case Active_Object
    Case "A"
        Call Move_Object_Left(A1, A2, A3, A4)
    Case "B"
        Call Move_Object_Left(B1, B2, B3, B4)
    Case "C"
        Call Move_Object_Left(C1, C2, C3, C4)
    Case "D"
        Call Move_Object_Left(D1, D2, D3, D4)
    Case "E"
        Call Move_Object_Left(E1, E2, E3, E4)
    Case "F"
        Call Move_Object_Left(F1, F2, F3, F4)
    Case "G"
        Call Move_Object_Left(G1, G2, G3, G4)
    End Select
End Sub

Private Sub Move_Any_Object_Right()
    Select Case Active_Object
    Case "A"
        Call Move_Object_Right(A1, A2, A3, A4)
    Case "B"
        Call Move_Object_Right(B1, B2, B3, B4)
    Case "C"
        Call Move_Object_Right(C1, C2, C3, C4)
    Case "D"
        Call Move_Object_Right(D1, D2, D3, D4)
    Case "E"
        Call Move_Object_Right(E1, E2, E3, E4)
    Case "F"
        Call Move_Object_Right(F1, F2, F3, F4)
    Case "G"
        Call Move_Object_Right(G1, G2, G3, G4)
    End Select
End Sub

Private Sub Move_Any_Object_Down()
    Select Case Active_Object
    Case "A"
        Call Move_Object_Down(A1, A2, A3, A4)
    Case "B"
        Call Move_Object_Down(B1, B2, B3, B4)
    Case "C"
        Call Move_Object_Down(C1, C2, C3, C4)
    Case "D"
        Call Move_Object_Down(D1, D2, D3, D4)
    Case "E"
        Call Move_Object_Down(E1, E2, E3, E4)
    Case "F"
        Call Move_Object_Down(F1, F2, F3, F4)
    Case "G"
        Call Move_Object_Down(G1, G2, G3, G4)
    End Select
End Sub

Private Sub Move_Object_Left(Obj1 As PictureBox, Obj2 As PictureBox, Obj3 As PictureBox, Obj4 As PictureBox)
    If UpdateMatricaLeft(Obj1, Obj2, Obj3, Obj4) Then
        Obj1.Left = Obj1.Left - x
        Obj2.Left = Obj2.Left - x
        Obj3.Left = Obj3.Left - x
        Obj4.Left = Obj4.Left - x
    End If
End Sub

Private Sub Move_Object_Right(Obj1 As PictureBox, Obj2 As PictureBox, Obj3 As PictureBox, Obj4 As PictureBox)
    If UpdateMatricaRight(Obj1, Obj2, Obj3, Obj4) Then
        Obj1.Left = Obj1.Left + x
        Obj2.Left = Obj2.Left + x
        Obj3.Left = Obj3.Left + x
        Obj4.Left = Obj4.Left + x
    End If
End Sub

Private Sub Move_Object_Down(Obj1 As PictureBox, Obj2 As PictureBox, Obj3 As PictureBox, Obj4 As PictureBox)
    If UpdateMatricaDown(Obj1, Obj2, Obj3, Obj4) Then
        Obj1.Top = Obj1.Top + x
        Obj2.Top = Obj2.Top + x
        Obj3.Top = Obj3.Top + x
        Obj4.Top = Obj4.Top + x
    End If
End Sub

'checks if object can be moved left
Private Function UpdateMatricaLeft(Obj1 As PictureBox, Obj2 As PictureBox, Obj3 As PictureBox, Obj4 As PictureBox) As Boolean
On Error GoTo ErrorHandler
    If ((Obj1.Left / x) > 0) _
    And (Matrica((Obj1.Left - x) / x, (Obj1.Top) / x) = False _
    And ((Obj2.Left / x) > 0) _
    And Matrica((Obj2.Left - x) / x, (Obj2.Top) / x) = False) _
    And ((Obj3.Left / x) > 0) _
    And Matrica((Obj3.Left - x) / x, (Obj3.Top) / x) = False _
    And ((Obj4.Left / x) > 0) _
    And Matrica((Obj4.Left - x) / x, (Obj4.Top) / x) = False Then
        UpdateMatricaLeft = True
    Else
ErrorHandler:
        UpdateMatricaLeft = False
    End If
End Function

'checks if object can be moved right
Private Function UpdateMatricaRight(Obj1 As PictureBox, Obj2 As PictureBox, Obj3 As PictureBox, Obj4 As PictureBox) As Boolean
On Error GoTo ErrorHandler
    If (((Obj1.Left) / x) < 11) _
    And (Matrica((Obj1.Left + x) / x, (Obj1.Top) / x) = False _
    And (((Obj2.Left) / x) < 11) _
    And Matrica((Obj2.Left + x) / x, (Obj2.Top) / x) = False) _
    And (((Obj3.Left) / x) < 11) _
    And Matrica((Obj3.Left + x) / x, (Obj3.Top) / x) = False _
    And (((Obj4.Left) / x) < 11) _
    And Matrica((Obj4.Left + x) / x, (Obj4.Top) / x) = False Then
        UpdateMatricaRight = True
    Else
ErrorHandler:
        UpdateMatricaRight = False
    End If
End Function

'checks if object can be moved down
Private Function UpdateMatricaDown(Obj1 As PictureBox, Obj2 As PictureBox, Obj3 As PictureBox, Obj4 As PictureBox) As Boolean
On Error GoTo ErrorHandler
    If (((Obj1.Top) / x) < 23) _
    And (Matrica((Obj1.Left) / x, (Obj1.Top + x) / x) = False _
    And (((Obj2.Top) / x) < 23) _
    And Matrica((Obj2.Left) / x, (Obj2.Top + x) / x) = False) _
    And (((Obj3.Top) / x) < 23) _
    And Matrica((Obj3.Left) / x, (Obj3.Top + x) / x) = False _
    And (((Obj4.Top) / x) < 23) _
    And Matrica((Obj4.Left) / x, (Obj4.Top + x) / x) = False Then
        UpdateMatricaDown = True
    Else
ErrorHandler:
        UpdateMatricaDown = False
    End If
End Function

Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    Dim Brojac As Integer
    
    Brojac = 0
    'sets background pictureboxes to initial positions
    For i = 0 To 23
        For j = 0 To 11
            Pic(Brojac).Left = j * x
            Pic(Brojac).Top = i * x
            Brojac = Brojac + 1
        Next j
    Next i
End Sub

Private Sub help_Click(Index As Integer)
    frmAbout.Show vbModal
End Sub

Private Sub new_Click(Index As Integer)
    Call Initialise
    Call Generate_New_Object
End Sub

Private Sub Timer1_Timer()
    If GameRunnig Then
        lblStatus.Caption = "Rows: " & NumberOfRows
        Call Move_Any_Object_Down
        Call GameOver
        Call CheckIsRowFull
        
        If Timer1.Interval > 100 Then
            Timer1.Interval = Speed - 4 * NumberOfRows
        End If
        Select Case Active_Object
        Case "A"
            If (UpdateMatricaDown(A1, A2, A3, A4) = False) Then
                Matrica(A1.Left / x, A1.Top / x) = True
                Matrica(A2.Left / x, A2.Top / x) = True
                Matrica(A3.Left / x, A3.Top / x) = True
                Matrica(A4.Left / x, A4.Top / x) = True
                
                Call HideObjects(A1, A2, A3, A4)
                Call DisplayFilledFields(A1, A2, A3, A4)
                Call Generate_New_Object
            End If
        Case "B"
            If (UpdateMatricaDown(B1, B2, B3, B4) = False) Then
                Matrica(B1.Left / x, B1.Top / x) = True
                Matrica(B2.Left / x, B2.Top / x) = True
                Matrica(B3.Left / x, B3.Top / x) = True
                Matrica(B4.Left / x, B4.Top / x) = True
                Call HideObjects(B1, B2, B3, B4)
                Call DisplayFilledFields(B1, B2, B3, B4)
                Call Generate_New_Object
            End If
        Case "C"
            If (UpdateMatricaDown(C1, C2, C3, C4) = False) Then
                Matrica(C1.Left / x, C1.Top / x) = True
                Matrica(C2.Left / x, C2.Top / x) = True
                Matrica(C3.Left / x, C3.Top / x) = True
                Matrica(C4.Left / x, C4.Top / x) = True
                Call HideObjects(C1, C2, C3, C4)
                Call DisplayFilledFields(C1, C2, C3, C4)
                Call Generate_New_Object
            End If
        Case "D"
            If (UpdateMatricaDown(D1, D2, D3, D4) = False) Then
                Matrica(D1.Left / x, D1.Top / x) = True
                Matrica(D2.Left / x, D2.Top / x) = True
                Matrica(D3.Left / x, D3.Top / x) = True
                Matrica(D4.Left / x, D4.Top / x) = True
                Call HideObjects(D1, D2, D3, D4)
                Call DisplayFilledFields(D1, D2, D3, D4)
                Call Generate_New_Object
            End If
        Case "E"
            If (UpdateMatricaDown(E1, E2, E3, E4) = False) Then
                Matrica(E1.Left / x, E1.Top / x) = True
                Matrica(E2.Left / x, E2.Top / x) = True
                Matrica(E3.Left / x, E3.Top / x) = True
                Matrica(E4.Left / x, E4.Top / x) = True
                Call HideObjects(E1, E2, E3, E4)
                Call DisplayFilledFields(E1, E2, E3, E4)
                Call Generate_New_Object
            End If
        Case "F"
            If (UpdateMatricaDown(F1, F2, F3, F4) = False) Then
                Matrica(F1.Left / x, F1.Top / x) = True
                Matrica(F2.Left / x, F2.Top / x) = True
                Matrica(F3.Left / x, F3.Top / x) = True
                Matrica(F4.Left / x, F4.Top / x) = True
                Call HideObjects(F1, F2, F3, F4)
                Call DisplayFilledFields(F1, F2, F3, F4)
                Call Generate_New_Object
            End If
        Case "G"
            If (UpdateMatricaDown(G1, G2, G3, G4) = False) Then
                Matrica(G1.Left / x, G1.Top / x) = True
                Matrica(G2.Left / x, G2.Top / x) = True
                Matrica(G3.Left / x, G3.Top / x) = True
                Matrica(G4.Left / x, G4.Top / x) = True
                Call HideObjects(G1, G2, G3, G4)
                Call DisplayFilledFields(G1, G2, G3, G4)
                Call Generate_New_Object
            End If
        End Select
    End If
End Sub

'this function is called each time when is neccesary
'to throw new object
Private Sub Generate_New_Object()
    Dim i As Integer
    Randomize
    i = Int(7 * Rnd)
    Call Restart_Objects
    
    Select Case i
    Case 0
        A1.Visible = True
        A2.Visible = True
        A3.Visible = True
        A4.Visible = True
        Active_Object = "A"
    Case 1
        B1.Visible = True
        B2.Visible = True
        B3.Visible = True
        B4.Visible = True
        Active_Object = "B"
    Case 2
        C1.Visible = True
        C2.Visible = True
        C3.Visible = True
        C4.Visible = True
        Active_Object = "C"
    Case 3
        D1.Visible = True
        D2.Visible = True
        D3.Visible = True
        D4.Visible = True
        Active_Object = "D"
    Case 4
        E1.Visible = True
        E2.Visible = True
        E3.Visible = True
        E4.Visible = True
        Active_Object = "E"
    Case 5
        F1.Visible = True
        F2.Visible = True
        F3.Visible = True
        F4.Visible = True
        Active_Object = "F"
    Case 6
        G1.Visible = True
        G2.Visible = True
        G3.Visible = True
        G4.Visible = True
        Active_Object = "G"
    End Select

End Sub

'initial position of objects
Private Sub Restart_Objects()
    A1.Top = x
    A1.Left = 6 * x
    A2.Top = 2 * x
    A2.Left = 6 * x
    A3.Top = 3 * x
    A3.Left = 6 * x
    A4.Top = 2 * x
    A4.Left = 7 * x
    B1.Top = 2 * x
    B1.Left = 6 * x
    B2.Top = 2 * x
    B2.Left = 7 * x
    B3.Top = x
    B3.Left = 7 * x
    B4.Top = x
    B4.Left = 8 * x
    C1.Top = x
    C1.Left = 5 * x
    C2.Top = x
    C2.Left = 6 * x
    C3.Top = x
    C3.Left = 7 * x
    C4.Top = x
    C4.Left = 8 * x
    D1.Top = x
    D1.Left = 5 * x
    D2.Top = x
    D2.Left = 6 * x
    D3.Top = 2 * x
    D3.Left = 6 * x
    D4.Top = 2 * x
    D4.Left = 7 * x
    E1.Top = x
    E1.Left = 5 * x
    E2.Top = 2 * x
    E2.Left = 5 * x
    E3.Top = 3 * x
    E3.Left = 5 * x
    E4.Top = x
    E4.Left = 6 * x
    F1.Top = x
    F1.Left = 6 * x
    F2.Top = 2 * x
    F2.Left = 6 * x
    F3.Top = 3 * x
    F3.Left = 6 * x
    F4.Top = x
    F4.Left = 5 * x
    G1.Top = x
    G1.Left = 5 * x
    G2.Top = x
    G2.Left = 6 * x
    G3.Top = 2 * x
    G3.Left = 5 * x
    G4.Top = 2 * x
    G4.Left = 6 * x
    Object_A_Position = 1
    Object_B_Position = 1
    Object_C_Position = 1
    Object_D_Position = 1
    Object_E_Position = 1
    Object_F_Position = 1
    Object_G_Position = 1
End Sub

'when object touches ground or other object
'that place gets new color
Private Sub DisplayFilledFields(Obj1 As PictureBox, Obj2 As PictureBox, Obj3 As PictureBox, Obj4 As PictureBox)
    Dim i As Integer
    For i = 0 To NumberOfFields
        If ((Pic(i).Left = Obj1.Left) And (Pic(i).Top) = Obj1.Top) Then Pic(i).Visible = True
        If ((Pic(i).Left = Obj2.Left) And (Pic(i).Top) = Obj2.Top) Then Pic(i).Visible = True
        If ((Pic(i).Left = Obj3.Left) And (Pic(i).Top) = Obj3.Top) Then Pic(i).Visible = True
        If ((Pic(i).Left = Obj4.Left) And (Pic(i).Top) = Obj4.Top) Then Pic(i).Visible = True
    Next i
End Sub

Private Sub HideObjects(Obj1 As PictureBox, Obj2 As PictureBox, Obj3 As PictureBox, Obj4 As PictureBox)
    Obj1.Visible = False
    Obj2.Visible = False
    Obj3.Visible = False
    Obj4.Visible = False
End Sub

Private Sub GameOver()
    Dim i As Integer
    Dim j As Integer
    
    'checks if object has reach top of screen
    For i = 0 To 11
        For j = 0 To 2
            If Matrica(i, j) = True Then
                 GameRunnig = False
                 DisplayHighScores = True
                 lblStatus = "  GAME OVER!!!   *** Rows: " & NumberOfRows & " ***"
                 Timer2.Enabled = True
            End If
        Next j
    Next i
End Sub

Private Sub CheckIsRowFull()
    Dim i As Integer
    Dim j As Integer
    Dim m As Integer
    Dim n As Integer
    
    For i = 0 To 23
        MatricaFull(i) = True
    Next i
        
    'checks is any row full
    For i = 23 To 0 Step -1
        For j = 11 To 0 Step -1
            If Matrica(j, i) = False Then
                MatricaFull(i) = False
            End If
        Next j
    Next i

    'if row is full, it's been removed from matrix
    For i = 0 To 23
        If MatricaFull(i) = True Then
            NumberOfRows = NumberOfRows + 1
            For m = i To 1 Step -1
                For n = 0 To 11
                    Matrica(n, m) = Matrica(n, m - 1)
                Next n
            Next m
        End If
    Next i

    'in background of playing screen is 288 (12*24) PictureBoxes
    'which becomes visible when object falls down
    'so it's neccesary to make them invisible when
    'row is been removed
    For i = 23 To 0 Step -1
        For j = 11 To 0 Step -1
            If Matrica(j, i) = False Then
                Pic(i * 12 + j).Visible = False
            Else
                Pic(i * 12 + j).Visible = True
            End If
        Next j
    Next i
End Sub

Private Sub Timer2_Timer()
    If lblStatus.ForeColor = vbBlack Then
        lblStatus.ForeColor = vbGreen
    Else
        lblStatus.ForeColor = vbBlack
    End If
    
    If DisplayHighScores Then frmHallOfFame.Show vbModal
End Sub
