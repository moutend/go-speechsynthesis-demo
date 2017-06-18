# Using SpeechSynthesizer API (WinRT) from Go

[![MIT License](http://img.shields.io/badge/license-MIT-blue.svg?style=flat-square)][license]

[license]: https://github.com/moutend/go-wca/blob/master/LICENSE

This example demonstrates:

- List the available voices.
- Generate a WAV audio according text file.
- Generate a WAV audio according SSML file.

## Prerequisites

This example supports only Windows 10 64 bit edition.

# Usage

Download this repository as zip or clone by git. Then open this repository with Command prompt or Powershell.

To list the available voices, run `speech.exe` without specifying any flags.

```console
C:\> ./speech
Available voices
================

1. (en-US) Microsoft David Mobile male
2. (en-US) Microsoft Zira Mobile female
3. (ja-JP) Microsoft Ayumi Mobile female
4. (ja-JP) Microsoft Haruka Mobile female
5. (en-US) Microsoft Mark Mobile male
6. (ja-JP) Microsoft Ichiro Mobile male
```

To generate a wav audio, specify `*.txt` file or `*.ssml` file with `-i` flag. After success, the wav audio file named `voice.wav` will be generated.

```console
C:\> ./speech -i hello.txt
```

## Build executable

```console
go build
```

That's it. If you want to build the executable on a non-Windows environment, please export the environment variable `GOOS='windows'` at first.

## Build `speechsynthesizer.dll`

The project file of the `speechsynthesizer.dll` is [here](https://github.com/moutend/speechsynthesizer).

## Contributing

1. Fork this repository
1. Create a feature branch
1. Add changes
1. Run `go fmt`
1. Commit your changes
1. Open a new Pull Request

## Author

[Yoshiyuki Koyanagi](https://github.com/moutend)
