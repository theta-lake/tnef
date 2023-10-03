// Package tnef extracts the body and attachments from Microsoft TNEF files.
package tnef

import (
	"bytes"
	"errors"
	"io/ioutil"
	"strings"
	//"unicode/utf8"
	"fmt"
	"regexp"
	// "encoding/hex"
)

const (
	tnefSignature = 0x223e9f78
	//lvlMessage    = 0x01
	lvlAttachment = 0x02
)

// These can be used to figure out the type of attribute
// an object is
const (
	ATTOWNER                   = 0x0000 // Owner
	ATTSENTFOR                 = 0x0001 // Sent For
	ATTDELEGATE                = 0x0002 // Delegate
	ATTDATESTART               = 0x0006 // Date Start
	ATTDATEEND                 = 0x0007 // Date End
	ATTAIDOWNER                = 0x0008 // Owner Appointment ID
	ATTREQUESTRES              = 0x0009 // Response Requested.
	ATTFROM                    = 0x8000 // From
	ATTSUBJECT                 = 0x8004 // Subject
	ATTDATESENT                = 0x8005 // Date Sent
	ATTDATERECD                = 0x8006 // Date Received
	ATTMESSAGESTATUS           = 0x8007 // Message Status
	ATTMESSAGECLASS            = 0x8008 // Message Class
	ATTMESSAGEID               = 0x8009 // Message ID
	ATTPARENTID                = 0x800a // Parent ID
	ATTCONVERSATIONID          = 0x800b // Conversation ID
	ATTBODY                    = 0x800c // Body
	ATTPRIORITY                = 0x800d // Priority
	ATTATTACHDATA              = 0x800f // Attachment Data
	ATTATTACHTITLE             = 0x8010 // Attachment File Name
	ATTATTACHMETAFILE          = 0x8011 // Attachment Meta File
	ATTATTACHCREATEDATE        = 0x8012 // Attachment Creation Date
	ATTATTACHMODIFYDATE        = 0x8013 // Attachment Modification Date
	ATTDATEMODIFY              = 0x8020 // Date Modified
	ATTATTACHTRANSPORTFILENAME = 0x9001 // Attachment Transport Filename
	ATTATTACHRENDDATA          = 0x9002 // Attachment Rendering Data
	ATTMAPIPROPS               = 0x9003 // MAPI Properties
	ATTRECIPTABLE              = 0x9004 // Recipients
	ATTATTACHMENT              = 0x9005 // Attachment
	ATTTNEFVERSION             = 0x9006 // TNEF Version
	ATTOEMCODEPAGE             = 0x9007 // OEM Codepage
	ATTORIGNINALMESSAGECLASS   = 0x9008 // Original Message Class
)

type tnefObject struct {
	Level  int
	Name   int
	Type   int
	Data   []byte
	Length int
}

// Attachment contains standard attachments that are embedded
// within the TNEF file, with the name and data of the file extracted.
type Attachment struct {
	Title      string
	Data       []byte
	Properties MsgPropertyList
}

/**
 * get a mapi attribute
 * @param  {[type]} c *Data)        GetMapiAttribute(attrId int) (attr *MAPIAttribute [description]
 * @return {[type]}   [description]
 */
func (c *Attachment) GetMapiAttribute(attrId int) (attr *MsgPropertyValue) {
	if len(c.Properties.Values) > 0 {
		for _, a := range c.Properties.Values {
			if int(a.TagId) == attrId {
				attr = a
				break
			}
		}
	}
	return
}

// ErrNoMarker signals that the file did not start with the fixed TNEF marker,
// meaning it's not in the TNEF file format we recognize (e.g. it just has the
// .tnef extension, or a wrong MIME type).
var ErrNoMarker = errors.New("file did not begin with a TNEF marker")

type MsgPropertyValue struct {
	TagType uint16
	TagId   uint16

	PropNameSpace []byte
	PropIDType    uint32
	PropMap       []byte // depend by Prop Id Type

	Data      interface{}
	DataCount uint32 // the number of elements from Data
	DataType  string
}
type MsgPropertyList struct {
	Values []*MsgPropertyValue
}

// Data contains the various data from the extracted TNEF file.
type Data struct {
	Body         []byte
	BodyHTML     []byte
	Attachments  []*Attachment
	Attributes   []MAPIAttribute
	MessageClass []byte
}

/**
 * get a mapi attribute
 * @param  {[type]} c *Data)        GetMapiAttribute(attrId int) (attr *MAPIAttribute [description]
 * @return {[type]}   [description]
 */
func (c *Data) GetMapiAttribute(attrId int) (attr *MAPIAttribute) {
	if len(c.Attributes) > 0 {
		for _, a := range c.Attributes {
			if a.Name == attrId {
				attr = &a
				break
			}
		}
	}
	return
}

/**
 * check if the attachment has a reference in html as cid
 * @param  {[type]} a *Attachment)  IsMimeRelated( [description]
 * @return {[type]}   [description]
 */
func (c *Data) AttachmentIsMimeRelated(a *Attachment) bool {

	attContentIdAttr := a.GetMapiAttribute(MAPITagAttachContentId)
	if attContentIdAttr == nil {
		return false
	}

	cid := attContentIdAttr.Data.(string)

	if c.BodyHTML != nil && len(c.BodyHTML) > 0 && cid != "" {
		re := `('|")[\s\t\r\n]*cid[\s\t\r\n]*\:[\s\t\r\n]*` + regexp.QuoteMeta(cid) + `[\s\t\r\n]*('|")`
		matched, err := regexp.Match(re, c.BodyHTML)
		if err == nil && matched {
			return true
		}
	}

	return false
}

func (a *Attachment) addAttr(obj tnefObject) {
	switch obj.Name {
	case ATTATTACHTITLE:
		a.Title = strings.Replace(string(obj.Data), "\x00", "", -1)
	case ATTATTACHDATA:
		a.Data = obj.Data
	default:
		//fmt.Printf("ATT Flag: %x Value: %v\r\n\r\n", obj.Name, string(obj.Data))
	}
}

// DecodeFile is a utility function that reads the file into memory
// before calling the normal Decode function on the data.
func DecodeFile(path string) (*Data, error) {
	data, err := ioutil.ReadFile(path)
	if err != nil {
		return nil, err
	}

	return Decode(data)
}

// Decode will accept a stream of bytes in the TNEF format and extract the
// attachments and body into a Data object.
func Decode(data []byte) (*Data, error) {
	if len(data) < 4 || byteToInt(data[0:4]) != tnefSignature {
		return nil, ErrNoMarker
	}

	//key := binary.LittleEndian.Uint32(data[4:6])
	offset := 6
	var attachment *Attachment
	tnef := &Data{
		Attachments: []*Attachment{},
	}

	for {
		obj, ok := decodeTNEFObject(data[offset:])
		if !ok {
			break
		}
		offset += obj.Length

		if obj.Name == ATTOEMCODEPAGE {
			//fmt.Printf("CODE PAGE: %s\r\n", bytes.TrimRight(obj.Data, "\x00"))
		} else if obj.Name == ATTMESSAGECLASS {
			tnef.MessageClass = bytes.TrimRight(obj.Data, "\x00")
		} else if obj.Name == ATTATTACHRENDDATA {
			/* Each set of attachment attributes MUST begin with the attAttachRendData attribute, followed by any
			other attributes; attachment properties encoded in the attAttachment attribute SHOULD be last.

			attAttachRendData = AttachType AttachPosition RenderWidth RenderHeight DataFlags
			AttachType = AttachTypeFile / AttachTypeOle
			AttachTypeFile=%x01.00
			AttachTypeOle=%x02.00
			AttachPosition= INT32
			RenderWidth=INT16
			RenderHeight=INT16
			DataFlags = FileDataDefault / FileDataMacBinary
			FileDataDefault= %x00.00.00.00
			FileDataMacBinary=%x01.00.00.00
			*/
			attachment = new(Attachment)
			tnef.Attachments = append(tnef.Attachments, attachment)

		} else if obj.Level == lvlAttachment {
			/*
				AttachAttribute = attrLevelAttachment idAttachAttr Length Data Checksum
				AttachProps = attrLevelAttachment idAttachment Length Data Checksum
			*/

			if obj.Name == ATTATTACHMENT {
				/*
									MAPI ATTR ID: 3616 (0xe20), Type: 0x0003 -> PidTagAttachSize | value: 3285 (bytes)
									MAPI ATTR ID: 12289 (0x3001), TAG Type: 30 (0x001e) -> PidTagDisplayName (type: 0x001f) | value: image001.jpg (same as PidTagAttachLongFilename)
									MAPI ATTR ID: 14082 (0x3702), TAG Type: 258 (0x0102) -> PidTagAttachEncoding | value: empty!!?? ->  If the attachment is in MacBinary format, this property is set to
				"{0x2A,86,48,86,F7,14,03,0B,01}"; otherwise, it is unset.
									MAPI ATTR ID: 14083 (0x3703), TAG Type: 30 (0x001e) -> PidTagAttachExtension (type: 0x001e) | value: .jpg
									MAPI ATTR ID: 14085 (0x3705), TAG Type: 3 (0x0003) -> PidTagAttachMethod | value: 1
									MAPI ATTR ID: 14087 (0x3707), TAG Type: 30 (0x1e) -> PidTagAttachLongFilename (0x001F) | value: image001.jpg
									MAPI ATTR ID: 14091 (0x370b), TAG Type: 3 (0x0003) -> PidTagRenderingPosition | value: -1 (-1 e de fapt 0xffffff, decoded as signed) ->  0xFFFFFFFF indicates a hidden attachment that is not to be rendered in the main text
									MAPI ATTR ID: 14094 (0x370e), TAG Type: 30 (0x001e) -> PidTagAttachMimeTag | value: image/jpeg
									MAPI ATTR ID: 14098 (0x3712), TAG Type: 30 (0x1e) -> PidTagAttachContentId | value: image001.jpg@01D49162.DB2DC760
									MAPI ATTR ID: 14100 (0x3714), TAG Type: 3 (0x3) -> PidTagAttachFlags | value: 4 (4 means attRenderedInBody)
									MAPI ATTR ID: 32762 (0x7ffa), TAG Type: 3 (0x3) -> PidTagAttachmentLinkId| value: 0 (must be 0, if is not overwriten)
									MAPI ATTR ID: 32763 (0x7ffb), TAG Type: 64 (0x0040) ->	PidTagExceptionStartTime|value: 915151392000000000
									MAPI ATTR ID: 32764 (0x7ffc), TAG Type: 64 (0x40) -> PidTagExceptionEndTime | value: 915151392000000000
									MAPI ATTR ID: 32765 (0x7ffd), TAG Type: 3 (0x3) -> PidTagAttachmentFlags | value: 8
									MAPI ATTR ID: 32766 (0x7ffe), TAG Type: 11 (0xb) -> PidTagAttachmentHidden| value: true
									MAPI ATTR ID: 32767 (0x7fff), TAG Type: 11 (0xb) -> PidTagAttachmentContactPhoto | value: false
									MAPI ATTR ID: 3617 (0x0e21), TAG Type: 3 (0x3) -> PidTagAttachNumber | value: 956325
									MAPI ATTR ID: 4088 (0x0ff8), TAG Type: 258 (0x0102) -> PidTagMappingSignature | value: 28 78 81 160 198 126 89 69 167 247 18 51 167 63 155 237
									MAPI ATTR ID: 4090 (0x0ffa), TAG Type: 258 (0x0102) -> ??? | value: 28 78 81 160 198 126 89 69 167 247 18 51 167 63 155 237
									MAPI ATTR ID: 4094 (0xffe), TAG Type: 3 (0x3) -> PidTagObjectType | value: 7 (7 means Attachment object)
									MAPI ATTR ID: 13325 (0x340d), TAG Type: 3 (0x3) -> PidTagStoreSupportMask | value: 245710845 ( Indicates whether string properties within the .msg file
										are Unicode-encoded.)
									MAPI ATTR ID: 13327 (0x340f), TAG Type: 3 (0x3) -> ??? | value: 245710845
				*/

				var err error
				attachment.Properties, err = decodeMsgPropertyList(obj.Data)
				if err != nil {
					return nil, err
				}
				//fmt.Printf("%v / %x\r\n", obj.Name, obj.Name)
				//fmt.Printf("%v", obj.Data)
			} else {
				attachment.addAttr(obj)
			}

			//fmt.Printf("TNEF Attach Level Flag ID: %x Value: %v\r\n\r\n", obj.Name, string(obj.Data))
		} else if obj.Name == ATTMAPIPROPS {
			var err error
			tnef.Attributes, err = decodeMapi(obj.Data)
			if err != nil {
				return nil, err
			}

			// Get the body property if it's there
			for _, attr := range tnef.Attributes {
				switch attr.Name {
				case MAPIBody:
					tnef.Body = attr.Data
				case MAPIBodyHTML:
					tnef.BodyHTML = attr.Data
				default:
					//fmt.Printf("MAPI Flag: %x Value: %v\r\n\r\n", attr.Name, string(attr.Data))
				}
			}
		} else {
			//fmt.Printf("TNEF Flag: %x Value: %s\r\n\r\n", obj.Name, obj.Data)
		}
	}

	return tnef, nil
}

/**
 * MessageAttribute = attrLevelMessage idMessageAttr Length Data Checksum
 * MessageProps = attrLevelMessage idMsgProps Length Data Checksum
 */
func decodeTNEFObject(data []byte) (object tnefObject, ok bool) {
	if len(data) < 9 {
		return
	}
	offset := 0

	object.Level = byteToInt(data[offset : offset+1])
	offset++
	object.Name = byteToInt(data[offset : offset+2])
	offset += 2
	object.Type = byteToInt(data[offset : offset+2])
	offset += 2
	attLength := byteToInt(data[offset : offset+4])
	offset += 4
	if offset+attLength+2 > len(data) {
		return
	}
	object.Data = data[offset : offset+attLength]
	offset += attLength
	//checksum := byteToInt(data[offset : offset+2])
	offset += 2

	object.Length = offset
	ok = true
	return
}

/**
 *  MsgPropertyList = MsgPropertyCount *MsgPropertyValue
 *  MsgPropertyCount = UINT32
 *  MsgPropertyValue = MsgPropertyTag MsgPropertyData
 *
 * @param  {[type]} data []byte)       (MsgPropertyList [description]
 * @return {[type]}      [description]
 */
func decodeMsgPropertyList(data []byte) (MsgPropertyList, error) {

	list := MsgPropertyList{Values: []*MsgPropertyValue{}}

	// little endian reader
	leReader := LittleEndianReader{}
	/*
		fmt.Println("------------------------")
		fmt.Printf("\r\n\r\n%s\r\n\r\n",hex.Dump(data))
		fmt.Printf("\r\n\r\n%s\r\n\r\n",data)
		fmt.Println("------------------------")
	*/
	if len(data) < 4 {
		return list, fmt.Errorf("decodeMsgPropertyList: data too short")
	}

	//  MsgPropertyCount *MsgPropertyValue

	offset := 0
	// no of properties encoded
	//	countValues := leReader.Uint32(data[offset:offset+4])

	offset += 4

	//fmt.Printf("Count values: %v\r\n", countValues)

	//MsgPropertyValue = MsgPropertyTag MsgPropertyData

	for {
		v := MsgPropertyValue{}

		//MsgPropertyTag = MsgPropertyType MsgPropertyId [NamedPropSpec]
		v.TagType = leReader.Uint16(data[offset : offset+2]) // 2 bytes
		//fmt.Printf("\r\nTAG Type: %#x Dump:\r\n%v", v.TagType, hex.Dump(data[offset:offset+2]))
		offset += 2

		// tagId is MAPI Property
		v.TagId = leReader.Uint16(data[offset : offset+2]) // 2 bytes
		//fmt.Printf("\r\nTag ID: %#x Dump:\r\n%v", v.TagId, hex.Dump(data[offset:offset+2]))
		offset += 2

		if v.TagId >= 0x8000 {
			// has  NamedPropSpec; NamedPropSpec = PropNameSpace PropIDType PropMap

			v.PropNameSpace = data[offset : offset+16] // guid - 16 bytes
			offset += 16
			v.PropIDType = leReader.Uint32(data[offset : offset+4])
			offset += 4
			if v.PropIDType == 0x00000000 {
				// should be an uint32 value
				v.PropMap = data[offset : offset+4]
				offset += 4
			} else {
				// propIDType == 0x01000000	=> is PropMap is string (PropMapString)
				// PropMapString = UINT32 *UINT16 %x00.00 [PropMapPad]
				valueLength := int(leReader.Uint32(data[offset : offset+4]))
				offset += 4

				tmpStr, bytesRead := leReader.Utf16(data[offset:], valueLength)
				offset += bytesRead

				v.PropMap = []byte(tmpStr)

				padd := 4 - bytesRead%4

				if padd < 4 {
					offset += padd
				}
			}
		}

		v.DataCount = 1

		//startValueIdx := offset

		switch v.TagType {
		case 0x0001: //NULL
		case 0x0002: //Int16
			v.DataType = "int16"
			// int16  - 2 bytes + 2 padding
			v.Data = leReader.Int16(data[offset : offset+2])
			offset += 4
		case 0x1002: //TypeMVInt16
			v.DataCount = leReader.Uint32(data[offset : offset+4])
			offset += 4
			var tmp []int16
			// extract the v.DataCount values of int16
			for i := 0; i < int(v.DataCount); i++ {
				tmp = append(tmp, leReader.Int16(data[offset:offset+2]))
				offset += 2
			}

			// skip the padding if exists
			if r := int(v.DataCount * 2 % 4); r != 0 {
				offset += r
			}
			v.Data = tmp
			v.DataType = "int16"
		case 0x0003: //TypeInt32
			v.Data = leReader.Int32(data[offset : offset+4]) // has padd x00 at the end
			offset += 4
			v.DataType = "int32"
		case 0x1003: //TypeMVInt32
			tmp := []int32{}
			v.DataCount = leReader.Uint32(data[offset : offset+4])
			offset += 4
			for i := 0; i < int(v.DataCount); i++ {
				tmp = append(tmp, leReader.Int32(data[offset:offset+4]))
				offset += 4
			}
			v.Data = tmp
			v.DataType = "int32"
		case 0x0004: //TypeFlt32
			v.Data = leReader.Float32(data[offset : offset+4])
			offset += 4
			v.DataType = "float32"
		case 0x1004: //TypeMVFlt32
			tmp := []float32{}
			v.DataCount = leReader.Uint32(data[offset : offset+4])
			offset += 4
			for i := 0; i < int(v.DataCount); i++ {
				tmp = append(tmp, leReader.Float32(data[offset:offset+4]))
				offset += 4
			}
			v.Data = tmp
			v.DataType = "float32"
		case 0x0005: //TypeFlt64
			v.Data = leReader.Float64(data[offset : offset+8])
			offset += 8
			v.DataType = "float64"
		case 0x1005: //TypeMVFlt64
			tmp := []float64{}
			v.DataCount = leReader.Uint32(data[offset : offset+4])
			offset += 4
			for i := 0; i < int(v.DataCount); i++ {
				tmp = append(tmp, leReader.Float64(data[offset:offset+8]))
				offset += 8
			}
			v.Data = tmp
			v.DataType = "float64"
		case 0x0006: //TypeCurrency  Signed 64-bit
			v.Data = leReader.Int64(data[offset : offset+8]) // has padd x00 at the end
			offset += 8
			v.DataType = "int64"
		case 0x1006: //TypeMVCurrency  Signed 64-bit
			tmp := []int64{}
			v.DataCount = leReader.Uint32(data[offset : offset+4])
			offset += 4
			for i := 0; i < int(v.DataCount); i++ {
				tmp = append(tmp, leReader.Int64(data[offset:offset+8]))
				offset += 8
			}
			v.Data = tmp
			v.DataType = "int64"
		case 0x0007: //TypeAppTime
			v.Data = leReader.Float64(data[offset : offset+8]) // has padd x00 at the end
			offset += 8
			v.DataType = "float64"
		case 0x1007: //TypeMVAppTime
			tmp := []float64{}
			v.DataCount = leReader.Uint32(data[offset : offset+4])
			offset += 4
			for i := 0; i < int(v.DataCount); i++ {
				tmp = append(tmp, leReader.Float64(data[offset:offset+8]))
				offset += 8
			}
			v.Data = tmp
			v.DataType = "float64"
		case 0x000B: //TypeBoolean - 16 bits
			v.Data = leReader.Int16(data[offset:offset+4]) > 0 // has padd x00 at the end
			offset += 4
		case 0x000D: //TypeObject -> unicode
			noOfValues := leReader.Uint32(data[offset : offset+4]) // should be always 1
			offset += 4
			tmp := make([][]byte, noOfValues)
			for i := 0; i < int(noOfValues); i++ {
				bytesLength := int(leReader.Uint32(data[offset : offset+4]))

				//fmt.Printf("String data length: %v Extracted Value: %v", stringLength, hex.Dump(data[offset:offset+4]))
				offset += 4

				/*
					tmpRunes := []rune{}
					for j:=0; j < int(stringLength); j++ {
						r, size := utf8.DecodeRune(data[offset:])
						tmpRunes = append(tmpRunes, r)
						offset += size
					}
					tmp[i] = string(tmpRunes)
				*/
				tmp[i] = data[offset : offset+bytesLength]
				offset += bytesLength

				padd := 4 - bytesLength%4
				if padd < 4 {
					offset += padd
				}
			}

			v.Data = tmp[0]
			v.DataType = "object"
		case 0x0014: //TypeInt64
			v.Data = leReader.Int64(data[offset : offset+8]) // has padd x00 at the end
			offset += 8
			v.DataType = "int64"
		case 0x1014: //TypeMVInt64
			tmp := []int64{}
			v.DataCount = leReader.Uint32(data[offset : offset+4])
			offset += 4
			for i := 0; i < int(v.DataCount); i++ {
				tmp = append(tmp, leReader.Int64(data[offset:offset+8]))
				offset += 8
			}
			v.Data = tmp
			v.DataType = "int64"
		case 0x001E, 0x101E: //TypeString8, TypeMVString8 -  8-bit character string with terminating null character. - multibyte character set (MBCS):
			// ATTOEMCODEPAGE???
			noOfValues := leReader.Uint32(data[offset : offset+4]) // should be always 1
			offset += 4

			tmp := make([]string, noOfValues)
			for i := 0; i < int(noOfValues); i++ {
				stringLength := int(leReader.Uint32(data[offset : offset+4]))

				//fmt.Printf("String data length: %v Extracted Value: %v", stringLength, hex.Dump(data[offset:offset+4]))
				offset += 4
				/**
				 * try to read stringLength chars and than calculate the number of bytes read
				 */

				tmpStr := data[offset : offset+stringLength]
				offset += stringLength

				//fmt.Println("String: ", string(tmpStr))
				/*

					rdr := bytes.NewReader(data[offset:])
					tmpStr := []byte{}
					for i:= 0; i< int(stringLength);i++ {
						rune, runeBytesSize, err := rdr.ReadRune()
						if err != nil {
							// we should return an error too
							return list, fmt.Errorf("decodeMsgPropertyList: error decoding %#x data type: %s", v.TagType, err)
						}
						tmpStr = append(tmpStr, []byte(string(rune))...)
						offset += runeBytesSize
					}
				*/
				// reads a multiple of 4; the rest must be padd it with 0x00
				padd := 4 - len(tmpStr)%4
				if padd < 4 {
					offset += padd
				}
				tmp[i] = string(bytes.TrimRight(tmpStr, "\x00"))
			}

			if v.TagType == 0x001E {
				if len(tmp) == 1 {
					v.Data = tmp[0]
				} else {
					v.Data = ""
				}
			} else {
				v.Data = tmp
			}
			v.DataType = "string"
		case 0x001F, 0x101F:
			//TypeUnicode (unicode utf16 LE string), TypeMVUnicode (array of unicode utf16 LE string)  - UTF-16LE or variant character string with terminating 2-byte null character.
			noOfValues := leReader.Uint32(data[offset : offset+4])
			offset += 4

			tmp := make([]string, noOfValues)
			for i := 0; i < int(noOfValues); i++ {
				stringLength := int(leReader.Uint32(data[offset : offset+4])) // no of bytes to read
				//fmt.Printf("String data length: %v Extracted Value: %v", stringLength, hex.Dump(data[offset:offset+4]))
				offset += 4

				/**
				 * try to read stringLength chars and than calculate the number of bytes read
				 */

				tmpStr, bytesRead := leReader.Utf16(data[offset:], stringLength)
				offset += bytesRead

				// reads a multiple of 4; the rest must be padd it with 0x00
				padd := 4 - (bytesRead % 4)
				//fmt.Printf("\r\nString Length To Read: %v Runes: %v", stringLength, utf8.RuneCountInString(tmpStr))
				//fmt.Printf("\r\nString bytes read: %v", bytesRead)
				if padd < 4 {
					offset += padd
					//fmt.Printf("\r\nString padd added: %v",padd)
				}

				tmp[i] = strings.TrimRight(tmpStr, "\x00")
			}

			if v.TagType == 0x001F {
				if len(tmp) == 1 {
					v.Data = tmp[0]
				} else {
					v.Data = ""
				}
			} else {
				v.Data = tmp
			}
			v.DataType = "string"
		case 0x0040: //TypeSystime - FILETIME (a PtypTime value, as specified in [MS-OXCDATA] section 2.11.1)
			v.Data = leReader.Uint64(data[offset : offset+8])
			offset += 8
		case 0x1040: //TypeMVSystime
			noOfValues := leReader.Uint32(data[offset : offset+4])
			offset += 4
			v.DataCount = noOfValues
			tmp := make([]uint64, noOfValues)
			for i := 0; i < int(noOfValues); i++ {
				tmp[i] = leReader.Uint64(data[offset : offset+8])
				offset += 8
			}
			v.Data = tmp
		case 0x0048: //TypeCLSID -  OLE GUID - 16 bytes
			v.Data = leReader.String(data[offset : offset+16])
			v.DataCount = 1
			offset += 16
		case 0x1048: //TypeMVCLSID
			tmp := []string{}
			v.DataCount = leReader.Uint32(data[offset : offset+4])
			offset += 4
			for i := 0; i < int(v.DataCount); i++ {
				tmp = append(tmp, leReader.String(data[offset:offset+16]))
				offset += 16
			}
			v.Data = tmp
		case 0x0102, 0x1102: //TypeBinary, /TypeMVBinary
			noOfValues := leReader.Uint32(data[offset : offset+4]) // should be always 1
			offset += 4

			tmp := make([][]byte, noOfValues)

			for i := 0; i < int(noOfValues); i++ {
				binaryLength := leReader.Uint32(data[offset : offset+4])
				//fmt.Printf("Binary data length: %v Extracted Value: %v", binaryLength, hex.Dump(data[offset:offset+4]))
				offset += 4

				tmp[i] = data[offset : offset+int(binaryLength)]
				offset += int(binaryLength)

				// reads a multiple of 4; the rest must be padd it with 0x00
				padd := 4 - int(binaryLength)%4
				if padd < 4 {
					offset += padd
				}
			}

			if v.TagType == 0x0102 {
				if len(tmp) == 1 {
					v.Data = tmp[0]
				} else {
					v.Data = []byte{}
				}
			} else {
				v.Data = tmp
			}
			v.DataType = "binary"
		default:
			return list, fmt.Errorf("decodeMsgPropertyList: data type %#x is invalid", v.TagType)
		}

		//fmt.Printf("\r\n\r\nTTag ID : %#x\r\nTag Type: %#x,\r\nTag Data: %v\r\nExtracted value:\r\n%v\r\n", v.TagId, v.TagType, v.Data, hex.Dump(data[startValueIdx:offset]))

		list.Values = append(list.Values, &v)

		if offset >= len(data) {
			break
		}
	}

	return list, nil
}
