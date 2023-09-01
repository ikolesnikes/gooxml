package excel

import (
	"bytes"
	"encoding/xml"
)

// XML prolog (processing instruction) is being written to every produced XML
// file.
var xmlProlog = xml.ProcInst{
	Target: "xml",
	Inst:   []byte("version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\""),
}

// encodeTokens encodes the given slice of tokens using the given encoder.
func encodeTokens(tokens []xml.Token, enc *xml.Encoder) error {
	for _, t := range tokens {
		if err := enc.EncodeToken(t); err != nil {
			return err
		}
	}
	return nil
}

func encodePart(part xml.Marshaler) (*bytes.Buffer, error) {
	b := new(bytes.Buffer)
	enc := xml.NewEncoder(b)
	defer enc.Close()

	if err := enc.Encode(part); err != nil {
		return nil, err
	}
	return b, enc.Flush()
}
