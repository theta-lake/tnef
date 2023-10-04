With this library you can extract the body and attachments from Transport
Neutral Encapsulation Format (TNEF) files.

This repository is a fork of https://github.com/Teamwork/tnef with fixes for attachment names and integrates various
other fixes from other forks.

This work is based on https://github.com/koodaamo/tnefparse and
http://www.freeutils.net/source/jtnef/.

## Example usage

```go
package main

import (
    "github.com/teamwork/tnef"
    "os"
)

func main() {
    t, err := tnef.DecodeFile("./winmail.dat")
    if err != nil {
        return
    }
    wd, _ := os.Getwd()
    for _, a := range t.Attachments {
        _ = os.WriteFile(wd+"/"+a.Title, a.Data, 0777)
    }
    _ = os.WriteFile(wd+"/bodyHTML.html", t.BodyHTML, 0777)
    _ = os.WriteFile(wd+"/bodyPlain.html", t.Body, 0777)
}

```
