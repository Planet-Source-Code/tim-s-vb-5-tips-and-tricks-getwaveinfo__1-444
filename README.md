<div align="center">

## GetWaveInfo


</div>

### Description

If you've ever wondered how sound applications can show the kilohertz and samples per second

information about a waveform file (.WAV), the answer lies in the RIFF file format.

The RIFF file format is designed to be as generic as possible. It is used for waveform, AVI, palette,

and other information standards that may need to be mixed and used together. Generally speaking,

though, any file with a WAV extension will only contain waveform data.

RIFF provides information in chunks and subchunks. The header for each chunk describes the

length of the chunk and the type of data the chunk contains (WAVE, for instance, is the string

identifying a WAVE chunk).

The Wave subchunk is immediately followed by the WAVE Format Chunk. It is this small chunk

that defines the structure of the waveform data that will follow. It defines the format of the

waveform, the number of channels used (with 0 being mono, 1 being stereo), the sampling rate, the

kilohertz at which is was recorded, and the data block size. Of these, only mono/stereo and the

sampling rate are likely to be of interest unless you intend to write your own custom waveform

player.
 
### More Info
 
had originally defined all of the string chunk identifiers (RIFF, WAVE, and 'fmt ') as being strings

in our user-defined data type WavInfo. But as fate would have it, I kept getting 'Bad File Handle'

errors when I used the string data types with VB5.0. So I elected to use a rather lengthy binary

representation of the same information, which follows the BUG FIX comment. I suspect that it has

something to do with Unicode, but really don't care to chase it down.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tim's VB 5 tips and tricks](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tim-s-vb-5-tips-and-tricks.md)
**Level**          |Unknown
**User Rating**    |4.0 (4 globes from 1 user)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tim-s-vb-5-tips-and-tricks-getwaveinfo__1-444/archive/master.zip)

### API Declarations

```

     'Type n, Mono/Stereo, 8/16 bit sample
     'These constants are not used internally, and
     'can be safely deleted if you do not intend to use them
     Public Const WAVE_FORMAT_1M08 = &H1
     Public Const WAVE_FORMAT_1M16 = &H4
     Public Const WAVE_FORMAT_1S08 = &H2
     Public Const WAVE_FORMAT_1S16 = &H8
     Public Const WAVE_FORMAT_2M08 = &H10
     Public Const WAVE_FORMAT_2M16 = &H40
     Public Const WAVE_FORMAT_2S08 = &H20
     Public Const WAVE_FORMAT_2S16 = &H80
     Public Const WAVE_FORMAT_4M08 = &H100
     Public Const WAVE_FORMAT_4M16 = &H400
     Public Const WAVE_FORMAT_4S08 = &H200
     Public Const WAVE_FORMAT_4S16 = &H800
     'BUG FIX
     'Binary representations of strings
     Public Const RIFF_ID = 1179011410
     Public Const RIFF_WAVE = 1163280727
     Public Const RIFF_FMT = 544501094
     'Typical header of a simple RIFF WAVE file
     Public Type WAVInfo
       Riff_Format As Long
       chunk_size As Long
       ChunkID As Long
       fmt As Long
       Wave_Format As Integer
       Channels As Integer       '0 = mono, 1 = stereo
       SamplesPerSecond As Long
       AverageBytesPerSecond As Long  '11.025kHz, 22.05kHz, etc
       BlockAlign As Integer      'Size of blocks for low level playback
     End Type
```


### Source Code

```
Public Function GetWaveInfo(Byval filename As String, Byref w As WAVInfo) _
       As Boolean
       Dim ff As Integer
       ff = FreeFile
       On Error GoTo ehandler
       Open filename For Binary Access Read As #ff
       On Error GoTo ehandler_fo
       Get #ff, , w
       Close #ff
       On Error GoTo ehandler
       If w.Riff_Format = RIFF_ID And w.ChunkID = _
         RIFF_WAVE And w.fmt = RIFF_FMT Then
         GetWaveInfo = True
       Else
         GetWaveInfo = False
       End If
       Exit Function
     ehandler_fo:
       Close #ff
     ehandler:
       GetWaveInfo = False
     End Function
```

