package office

import (
	"strings"
	"testing"
)

func TestGeneratedCommentIDsPreserveLinks(t *testing.T) {
	a := `<root xmlns:w="` + wordXMLNamespace + `" xmlns:p="` + word14 + `" xmlns:x="` + word15 + `" xmlns:i="` + wordCID + `" xmlns:e="` + wordCEX + `"><w:comment w:id="0"><w:p p:paraId="AAA"><w:r><w:t>First</w:t></w:r></w:p></w:comment><w:comment w:id="1"><w:p p:paraId="BBB"><w:r><w:t>Reply</w:t></w:r></w:p></w:comment><x:commentEx x:paraId="AAA" x:done="0"/><x:commentEx x:paraId="BBB" x:paraIdParent="AAA"/><i:commentId i:paraId="AAA" i:durableId="CCC"/><i:commentId i:paraId="BBB" i:durableId="DDD"/><e:commentExtensible e:durableId="CCC"/><e:commentExtensible e:durableId="DDD"/></root>`
	b := strings.NewReplacer("AAA", "111", "BBB", "222", "CCC", "333", "DDD", "444").Replace(a)
	policy := XMLComparePolicy{GeneratedCommentIDs: true}
	result, err := CompareXML([]byte(a), []byte(b), policy)
	if err != nil || result["equal"] != true || result["byte_identical"] != false {
		t.Fatal(result, err)
	}
	result, err = CompareXML([]byte(a), []byte(b), XMLComparePolicy{})
	if err != nil || result["equal"] != false {
		t.Fatal("normalization must be opt-in", result, err)
	}
	for _, changed := range []string{
		strings.Replace(b, `paraIdParent="111"`, `paraIdParent="222"`, 1),
		strings.Replace(b, `e:durableId="333"`, `e:durableId="444"`, 1),
		strings.Replace(b, `x:done="0"`, `x:done="1"`, 1),
		strings.Replace(b, "Reply", "Changed", 1),
		strings.Replace(b, `paraIdParent="111"`, `paraIdParent="missing"`, 1),
		strings.Replace(b, `p:paraId="222"`, `p:paraId="111"`, 1),
	} {
		result, err = CompareXML([]byte(a), []byte(changed), policy)
		if err == nil && result["equal"] == true {
			t.Fatal("accepted changed comment semantics", changed)
		}
	}
}
