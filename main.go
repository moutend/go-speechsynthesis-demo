// +build windows
package main

import (
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"reflect"
	"strings"
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/moutend/go-wss"
)

type FilenameFlag struct {
	Value     string
	Extension string
}

func (f *FilenameFlag) Set(value string) (err error) {
	s := strings.Split(value, ".")
	f.Value = value
	f.Extension = s[len(s)-1]
	return
}

func (f *FilenameFlag) String() string {
	return f.Value
}

func main() {
	var err error
	if err = run(os.Args); err != nil {
		log.Fatal(err)
	}
	return
}

func run(args []string) (err error) {
	var inputFlag FilenameFlag
	var outputFlag FilenameFlag
	var voiceNumberFlag int
	var file []byte
	var stream []byte

	f := flag.NewFlagSet(args[0], flag.ExitOnError)
	f.Var(&inputFlag, "i", "input file name (*.txt or *.ssml)")
	f.Var(&outputFlag, "o", "output file name (default: voice.wav)")
	f.IntVar(&voiceNumberFlag, "n", 0, "voice number")
	f.Parse(args[1:])

	if inputFlag.Value == "" {
		err = ListVoices()
		return
	}
	if file, err = ioutil.ReadFile(inputFlag.Value); err != nil {
		return
	}
	if outputFlag.Value == "" {
		outputFlag.Value = "voice.wav"
	}
	if stream, err = CreateSpeechStream(voiceNumberFlag, inputFlag.Extension, string(file[:])); err != nil {
		return
	}
	return ioutil.WriteFile(outputFlag.Value, stream, 0644)
}

func CreateSpeechStream(voiceNumber int, format, text string) (stream []byte, err error) {
	var modSpeechSynthesizer *syscall.DLL
	if modSpeechSynthesizer, err = syscall.LoadDLL("speechsynthesizer.dll"); err != nil {
		return
	}

	var proc *syscall.Proc
	if format == "ssml" {
		proc, err = modSpeechSynthesizer.FindProc("SsmlToStream")
	} else {
		proc, err = modSpeechSynthesizer.FindProc("TextToStream")
	}
	if err != nil {
		return
	}

	var textHString ole.HString
	if textHString, err = ole.NewHString(text); err != nil {
		return
	}
	defer ole.DeleteHString(textHString)

	ole.RoInitialize(1)
	//defer ole.RoUninitialize()

	var hr uintptr
	var buf *wss.IBuffer
	hr, _, _ = proc.Call(
		uintptr(uint32(voiceNumber)),
		uintptr(unsafe.Pointer(textHString)),
		uintptr(unsafe.Pointer(&buf)))
	if hr != 0 {
		err = fmt.Errorf("unknown error %v", hr)
		return
	}
	defer buf.Release()

	var length uint32
	if err = buf.GetLength(&length); err != nil {
		return
	}

	// Cast Ibuffer to IBufferByteAccess.
	var unk *ole.IUnknown
	if buf.PutQueryInterface(ole.IID_IUnknown, &unk); err != nil {
		return
	}

	var bba *wss.IBufferByteAccess
	if err = unk.PutQueryInterface(wss.IID_IBufferByteAccess, &bba); err != nil {
		return
	}

	// Extract the backing array.
	var bufPtr *byte
	if err = bba.Buffer(&bufPtr); err != nil {
		return
	}

	// Convert native byte array to Go's byte slice.
	rawBufPtr := uintptr(unsafe.Pointer(bufPtr))
	sliceHeader := reflect.SliceHeader{Data: rawBufPtr, Len: int(length), Cap: int(length)}
	stream = *(*[]byte)(unsafe.Pointer(&sliceHeader))

	return
}

func ListVoices() (err error) {
	var modSpeechSynthesizer *syscall.DLL
	if modSpeechSynthesizer, err = syscall.LoadDLL("speechsynthesizer.dll"); err != nil {
		return
	}

	var proc *syscall.Proc
	if proc, err = modSpeechSynthesizer.FindProc("GetVoices"); err != nil {
		return
	}

	ole.RoInitialize(1)
	//defer ole.RoUninitialize()

	var vv *wss.IVectorView
	hr, _, _ := proc.Call(
		uintptr(unsafe.Pointer(&vv)))
	if hr != 0 {
		err = fmt.Errorf("unknown error %v", hr)
		return
	}
	defer vv.Release()

	var size uint16
	if err = vv.GetSize(&size); err != nil {
		return
	}

	var vi *wss.IVoiceInformation
	var name ole.HString
	var language ole.HString
	var gender wss.VoiceGender
	var genderString string

	fmt.Println("Available voices")
	fmt.Println("================\n")

	for i := 0; i < int(size); i++ {
		if err = vv.GetAt(uint16(i), &vi); err != nil {
			return
		}
		if err = vi.GetDisplayName(&name); err != nil {
			return
		}
		if err = vi.GetLanguage(&language); err != nil {
			return
		}
		if err = vi.GetGender(&gender); err != nil {
			return
		}
		if gender == wss.VoiceGender_Male {
			genderString = "male"
		} else {
			genderString = "female"
		}
		fmt.Printf("%d. (%v) %v %v\n", i+1, language.String(), name.String(), genderString)
		vi.Release()
	}
	return
}
