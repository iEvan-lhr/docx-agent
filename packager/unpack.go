package packager

import (
	"archive/zip"
	"bytes"
	"fmt"
	"path"
	"strings"

	"github.com/iEvan-lhr/docx-agent/common/constants"
	"github.com/iEvan-lhr/docx-agent/docx"
	"github.com/iEvan-lhr/docx-agent/internal"
	"github.com/iEvan-lhr/docx-agent/wml/ctypes"
)

// ReadFromZip reads files from a zip archive.
func ReadFromZip(content *[]byte) (map[string][]byte, error) {
	zipReader, err := zip.NewReader(bytes.NewReader(*content), int64(len(*content)))
	if err != nil {
		return nil, err
	}

	var (
		fileList = make(map[string][]byte, len(zipReader.File))
	)

	for _, f := range zipReader.File {

		fileName := strings.ReplaceAll(f.Name, "\\", "/")

		if fileList[fileName], err = internal.ReadFileFromZip(f); err != nil {
			return nil, err
		}
	}

	return fileList, nil
}

func Unpack(content *[]byte) (*docx.RootDoc, error) {

	rd := docx.NewRootDoc()

	fileIndex, err := ReadFromZip(content)
	if err != nil {
		return nil, err
	}

	// Load content type details
	ctBytes := fileIndex[constants.ConentTypeFileIdx]
	ct, err := LoadContentTypes(ctBytes)
	if err != nil {
		return nil, err
	}
	delete(fileIndex, constants.ConentTypeFileIdx)
	rd.ContentType = *ct

	rd.ImageCount = 0

	rootRelURI, err := GetRelsURI("")
	if err != nil {
		return nil, err
	}

	rootRelBytes := fileIndex[*rootRelURI]
	rootRelations, err := LoadRelationShips(*rootRelURI, rootRelBytes)
	if err != nil {
		return nil, err
	}
	delete(fileIndex, *rootRelURI)
	rd.RootRels = *rootRelations

	var docPath string

	for _, relation := range rootRelations.Relationships {
		switch relation.Type {
		case constants.OFFICE_DOC_TYPE:
			docPath = relation.Target
		}
	}

	if docPath == "" {
		return nil, fmt.Errorf("root officeDocument type not found")
	}

	docRelURI, err := GetRelsURI(docPath)
	if err != nil {
		return nil, err
	}

	// Load document
	docFile := fileIndex[docPath]
	docObj, err := docx.LoadDocXml(rd, docPath, docFile)
	if err != nil {
		return nil, err
	}
	delete(fileIndex, docPath)
	rd.Document = docObj

	// 在加载 document 后，初始化 Headers 和 Footers map
	rd.Document.Headers = make(map[string]*docx.Header)
	rd.Document.Footers = make(map[string]*docx.Footer)

	// Load Relationship details
	docRelFile := fileIndex[*docRelURI]
	docRelations, err := LoadRelationShips(*docRelURI, docRelFile)
	if err != nil {
		return nil, err
	}
	delete(fileIndex, *rootRelURI)
	rd.Document.DocRels = *docRelations

	wordDir := path.Dir(docPath)

	rd.DocStyles = &ctypes.Styles{}
	rID := 0
	for _, relation := range docRelations.Relationships {
		rID += 1
		switch relation.Type {
		case constants.StylesType:
			sFileName := relation.Target
			if sFileName == "" {
				continue
			}
			stylesPath := path.Join(wordDir, sFileName)

			//Load Styles
			stylesFile := fileIndex[stylesPath]
			stylesObj, err := docx.LoadStyles(stylesPath, stylesFile)
			if err != nil {
				return nil, err
			}
			delete(fileIndex, stylesPath)
			rd.DocStyles = stylesObj
		case constants.HeaderType:
			// 处理 Header
			headerFileName := relation.Target
			if headerFileName == "" {
				continue
			}
			headerPath := path.Join(wordDir, headerFileName)

			// 加载 Header
			headerFile := fileIndex[headerPath]
			headerObj, err := docx.LoadHeaderXml(rd, headerPath, headerFile)
			if err != nil {
				return nil, fmt.Errorf("failed to load header %s: %v", headerPath, err)
			}
			headerObj.ID = relation.ID
			rd.Document.Headers[relation.ID] = headerObj
			delete(fileIndex, headerPath)

		case constants.FooterType:
			// 处理 Footer
			footerFileName := relation.Target
			if footerFileName == "" {
				continue
			}
			footerPath := path.Join(wordDir, footerFileName)

			// 加载 Footer
			footerFile := fileIndex[footerPath]
			footerObj, err := docx.LoadFooterXml(rd, footerPath, footerFile)
			if err != nil {
				return nil, fmt.Errorf("failed to load footer %s: %v", footerPath, err)
			}
			footerObj.ID = relation.ID
			rd.Document.Footers[relation.ID] = footerObj
			delete(fileIndex, footerPath)
		}
	}

	rd.Document.RID = rID

	for fileName, fileContent := range fileIndex {
		if strings.HasPrefix(fileName, constants.MediaPath) {
			rd.ImageCount += 1
		}
		rd.FileMap.Store(fileName, fileContent)
	}

	return rd, nil
}
