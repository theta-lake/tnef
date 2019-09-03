package tnef

import (
	"encoding/binary"
	"bytes"
	"unicode/utf16"
//	"unicode/utf8"
	"strings"
)

func byteToInt(data []byte) int {
	var num int
	var n uint
	for _, b := range data {
		num += (int(b) << n)
		n += 8
	}
	return num
}

/*
func byteToUInt32(data []byte) uint32 {
	return binary.LittleEndian.Uint32(data)
}*/



type LittleEndianReader struct {

}

func (c *LittleEndianReader) String(b []byte) (string) {
	var v string
	buf := bytes.NewReader(b)
 	 binary.Read(buf, binary.LittleEndian, &v)
	return v
}

// Int
func (c *LittleEndianReader) Int(b []byte) (int) {
	var v int
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}
// UInt
func (c *LittleEndianReader) Uint(b []byte) (uint) {
	var v uint
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}

// Int = Int32
func (c *LittleEndianReader) Int32(b []byte) (int32) {
	var v int32
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}
// UInt = UInt32
func (c *LittleEndianReader) Uint32(b []byte) (uint32) {
	var v uint32
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}

// int64
func (c *LittleEndianReader) Int64(b []byte) (int64) {
	var v int64
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}
// uint64
func (c *LittleEndianReader) Uint64(b []byte) (uint64) {
	var v uint64
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}

func (c *LittleEndianReader) Int16(b []byte) (int16) {
	var v int16
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}

func (c *LittleEndianReader) Uint16(b []byte) (uint16) {
	var v uint16
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}

func (c *LittleEndianReader) Int8(b []byte) (int8) {
	var v int8
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}

func (c *LittleEndianReader) Uint8(b []byte) (uint8) {
	var v uint8
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}

func (c *LittleEndianReader) Float32(b []byte) (float32) {
	var v float32
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}

func (c *LittleEndianReader) Float64(b []byte) (float64) {
	var v float64
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}

func (c *LittleEndianReader) Boolean(b []byte) (bool) {
	var v bool
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}

func (c *LittleEndianReader) Utf16(content []byte, maxLengthToRead int) (convertedStringToUnicode string, bytesRead int) {
	tmp := []uint16{}

	bytesRead = 0
	for {
		tmp = append(tmp, binary.LittleEndian.Uint16(content[bytesRead:]))
		bytesRead += 2

		convertedStringToUnicode = string(utf16.Decode(tmp));

		if (len(content) <= bytesRead || len(convertedStringToUnicode) == maxLengthToRead) {
			break;
		}
	}
	return
}

// read utf or unicode ended with 2 0x00
func (c *LittleEndianReader) Utf16OrUnicode(content []byte, maxLengthToRead int) (convertedStringToUnicode string, bytesRead int) {
	tmp := []uint16{}

	bytesRead = 0
	last2Chars := 0
	for {
		tmp = append(tmp, binary.LittleEndian.Uint16(content[bytesRead:]))
		bytesRead += 2


		convertedStringToUnicode = string(utf16.Decode(tmp));

		if strings.HasSuffix(convertedStringToUnicode, "\x00") {
			last2Chars++
		} else {
			last2Chars = 0
		}

		// utf8.RuneCountInString(convertedStringToUnicode) -> number of chars
		// len(convertedStringToUnicode) - number of bytes

//		fmt.Printf("\r\nRead %v bytes from %v bytes; String Len: %v; Rune Count: %v", bytesRead, len(content), len(convertedStringToUnicode), utf8.RuneCountInString(convertedStringToUnicode))
//		fmt.Printf("\r\nString: %s", convertedStringToUnicode)


		if (len(content) <= bytesRead || len(convertedStringToUnicode) == maxLengthToRead || last2Chars >= 2) {
			if (last2Chars >= 2) {
				convertedStringToUnicode = strings.TrimRight(convertedStringToUnicode, "\x00")
			}
			break;
		}
	}
	return
}
