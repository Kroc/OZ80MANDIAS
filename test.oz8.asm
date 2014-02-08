//a file to test OZ80MANDIAS parsing and assembling

//test each of the Z80 mnemonics and parameter styles

//add with carry
 //adc A, r|$n|(HL|IX+$n|IY+$n)
 //adc HL, BC|DE|HL|SP
adc a, b
adc a, $FF
adc a, (hl)
adc a, (ix+$01)
adc hl, bc
