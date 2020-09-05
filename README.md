# mono-to-stereo

Takes a mono input and renders it as if it was an interleaved stereo input. Works on MS2109 capture
devices where the audio input is a 96khz mono stream but in actuality is a 48khz stereo stream with
the first left channel sample missing. In order to support this device better, accounting for this
missing first sample is done by default.

Original code based off of [Matthew van Eerde's loopback-capture](https://github.com/mvaneerde/blog/tree/master/loopback-capture)
project.

Run `mono-to-stereo.exe -?` for usage instructions.
