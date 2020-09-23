<div align="center">

## CRC32 Algorithm


</div>

### Description

Pure visual basic implementation of the CRC32 (Fletcher) checksum algorithm. Used for detection of data corruption. In short, this class provides an algorithm for computing a unique (sort of) numeric value that represents the composition of a file. For a faster/better method see Adler32 checksum. Note, Fletcher check insensitive to some single byte changes 0 <-> 255.
 
### More Info
 
CRC32(ByVal lngCRC32 As Long, ByRef bArrayIn() As Byte, ByVal dblLength As Double) As Long

Files size assumed to be less than 2 gig. When passing byte arrays be carefull of the file size. It is best to break the file into chunks and call CRC32 multiple times.


<span>             |<span>
---                |---
**Submitted On**   |2002-08-25 15:56:28
**By**             |[Ed Preston](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ed-preston.md)
**Level**          |Advanced
**User Rating**    |5.0 (35 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CRC32\_Algo1222478252002\.zip](https://github.com/Planet-Source-Code/ed-preston-crc32-algorithm__1-38270/archive/master.zip)








