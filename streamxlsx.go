// Package streamxlsx implements a streaming .xlsx (Excel spreadsheat) file writer.
//
// The focus is to easily generate tabular data files.
package streamxlsx

import (
	"archive/zip"
	"fmt"
	"io"
	"path/filepath"
)

type StreamXLSX struct {
	zip            *zip.Writer
	openSheet      *sheetEncoder
	finishedSheets []string
	// The stylesheet will be written on Close(). You generally won't want to
	// use this directly, but via `Format()`.
	Styles      *Stylesheet
	styleCache  map[string]int
	ImageLoader ImageLoader
	error       error // returned with Close()
	// Image tracking
	allImages    []imageRef   // all images across all sheets
	sheetImages  [][]imageRef // images per sheet (indexed by sheet number)
	currentSheet int          // current sheet index (0-based)
	imageFormats []string     // collected image formats (for content types)
	// Sheet-level relations tracking (for hyperlinks and drawings)
	sheetRelations [][]relationship // relations per sheet (for writing in Close)
}

// New creates a new file. Do Close() it afterwards. No need to check every
// write for errors, Close() will return the last error.
//
// A StreamXLSX is not safe to use with multiple go routines at the same time.
func New(w io.Writer) *StreamXLSX {
	s := &StreamXLSX{
		zip:            zip.NewWriter(w),
		Styles:         &Stylesheet{},
		styleCache:     map[string]int{},
		sheetImages:    make([][]imageRef, 0),
		sheetRelations: make([][]relationship, 0),
		currentSheet:   0,
	}

	// empty style. Not 100% it's needed
	styleID := s.Styles.GetCellStyleID(Xf{})
	s.Styles.GetCellID(Xf{XfID: &styleID})

	s.writeRelations()
	return s
}

// Write a row to the current sheet.
// No values is a valid (empty) row, and not every row needs to have the same
// number of elements.
//
// Supported cell datatypes:
//
//	all ints and uints, floats, string, bool
//
// Additional special cases:
//
//	[]byte: will be base64 encoded
//	time.Time: handled, but you need to Format() it. For example: s.Format("mm-dd-yy", aTimeTime)
//	Hyperlink{}: will make the cell a hyperlink
//	Cell{}: if you want to set everything manually
//
// See Format() to apply number formatting to cells.
// See WriteRowOpts() to set row options like the row height.
func (s *StreamXLSX) WriteRow(vs ...interface{}) error {
	return s.WriteRowOpts(RowOpts{}, vs...)
}

// WriteRowOpts writes a row to the current sheet with the given options.
// The vs arguments are the same as for WriteRow().
//
// Example:
//
//	s.WriteRowOpts(RowOpts{Height: 50}, "label", Image{Name: "big.png", Height: 45})
func (s *StreamXLSX) WriteRowOpts(opts RowOpts, vs ...interface{}) error {
	if s.error != nil {
		return s.error
	}

	sh, err := s.sheet()
	if err != nil {
		s.error = err
		return err
	}
	if err := sh.writeRow(opts, vs...); err != nil {
		s.error = err
		return err
	}
	return nil
}

// WriteSheet closes the currenly open sheet, with the given title.
// The process is you first do all the `WriteRow()`s for a sheet, followed by
// its WriteSheet().  There is always an open sheet. You don't have to close
// the final sheet, but it'll give you a boring name ("sheet N").
//
// Note that Excel doesn't like titles over 31 characters and might give an
// error on opening. Openoffice and Google sheets don't seem to care much.
func (s *StreamXLSX) WriteSheet(title string) error {
	if s.error != nil {
		return s.error
	}

	// make sure there is a sheet open
	if _, err := s.sheet(); err != nil {
		s.error = err
		return err
	}
	// Collect images and relations from this sheet before closing
	if s.openSheet != nil {
		sheetIndex := s.openSheet.sheetIndex
		// Ensure we have a slot for this sheet's images
		for len(s.sheetImages) <= sheetIndex {
			s.sheetImages = append(s.sheetImages, []imageRef{})
		}
		if len(s.openSheet.images) > 0 {
			s.sheetImages[sheetIndex] = s.openSheet.images
			s.allImages = append(s.allImages, s.openSheet.images...)
			s.openSheet.drawingRelID = fmt.Sprintf("drawId%d", len(s.openSheet.relations)+1)
		}
		for len(s.sheetRelations) <= sheetIndex {
			s.sheetRelations = append(s.sheetRelations, []relationship{})
		}
		s.sheetRelations[sheetIndex] = append(s.sheetRelations[sheetIndex], s.openSheet.relations...)
	}
	s.openSheet.Close()
	s.openSheet = nil
	s.finishedSheets = append(s.finishedSheets, title)
	return nil
}

// Adds a number format to a cell. Examples of formats are "0.00", "0%", ...
// This is used to wrap a value in a WriteRow().
func (s *StreamXLSX) Format(code string, cell interface{}) Cell {
	if xfID, ok := s.styleCache[code]; ok {
		c, err := applyStyle(xfID, cell)
		if err != nil {
			s.error = err
		}
		return c
	}

	numFmtID := s.Styles.GetNumFmtID(code)
	cellStyleID := s.Styles.GetCellStyleID(Xf{})
	styleFx := Xf{
		NumFmtID:          numFmtID,
		ApplyNumberFormat: 1,
		XfID:              &cellStyleID,
	}
	xfID := s.Styles.GetCellID(styleFx)
	s.styleCache[code] = xfID
	c, err := applyStyle(xfID, cell)
	if err != nil {
		s.error = err
	}
	return c
}

// Adds a hyperlink in a cell. You can use these as a value in WriteRow().
// (implementation detail: parts of the hyperlink datastructure is
// only written when closing a sheet, so they are buffered)
type Hyperlink struct {
	URL, Title, Tooltip string
}

// Finish writing the spreadsheet.
func (s *StreamXLSX) Close() error {
	if s.error != nil {
		return s.error
	}

	if len(s.finishedSheets) == 0 {
		// there seems to be a requirement of at least 1 sheet
		if _, err := s.sheet(); err != nil {
			return err
		}
	}
	if s.openSheet != nil {
		// Collect images and relations from the open sheet before closing it
		sheetIndex := s.openSheet.sheetIndex
		// Ensure we have a slot for this sheet's images
		for len(s.sheetImages) <= sheetIndex {
			s.sheetImages = append(s.sheetImages, []imageRef{})
		}
		if len(s.openSheet.images) > 0 {
			s.sheetImages[sheetIndex] = s.openSheet.images
			s.allImages = append(s.allImages, s.openSheet.images...)
			s.openSheet.drawingRelID = fmt.Sprintf("drawId%d", len(s.openSheet.relations)+1)
		}
		for len(s.sheetRelations) <= sheetIndex {
			s.sheetRelations = append(s.sheetRelations, []relationship{})
		}
		s.sheetRelations[sheetIndex] = append(s.sheetRelations[sheetIndex], s.openSheet.relations...)

		s.openSheet.Close()
		s.finishedSheets = append(s.finishedSheets, fmt.Sprintf("sheet %d", len(s.finishedSheets)+1))
		s.openSheet = nil
	}

	// Write sheet relations (including drawing relationships for images)
	if err := s.writeAllSheetRelations(); err != nil {
		return err
	}

	// Write images and drawings if we have any
	if s.ImageLoader != nil && len(s.allImages) > 0 {
		var err error
		s.imageFormats, err = s.writeImagesAndDrawings()
		if err != nil {
			return err
		}
	}

	if err := s.writeWorkbook(); err != nil {
		return err
	}
	if err := s.writeStylesheet(); err != nil {
		return err
	}
	if err := s.writeSharedStrings(); err != nil {
		return err
	}
	if err := s.writeWorkbookRelations(); err != nil {
		return err
	}
	if err := s.writeContentTypes(); err != nil {
		return err
	}
	if err := s.zip.Close(); err != nil {
		return err
	}
	return nil
}

func (s *StreamXLSX) sheet() (*sheetEncoder, error) {
	if s.openSheet != nil {
		return s.openSheet, nil
	}
	sheetIndex := len(s.finishedSheets)
	s.currentSheet = sheetIndex
	// Ensure we have a slot for this sheet's images
	for len(s.sheetImages) <= sheetIndex {
		s.sheetImages = append(s.sheetImages, []imageRef{})
	}
	filename := fmt.Sprintf("xl/worksheets/sheet%d.xml", sheetIndex+1)
	fh, err := s.zip.Create(filename) // no need to close!
	if err != nil {
		return nil, err
	}

	enc, err := newSheetEncoder(fh, sheetIndex)
	if err != nil {
		return nil, err
	}

	s.openSheet = enc
	return s.openSheet, nil
}

func (s *StreamXLSX) writeWorkbook() error {
	filename := "xl/workbook.xml"
	fh, err := s.zip.Create(filename)
	if err != nil {
		return err
	}
	return writeWorkbook(fh, s.finishedSheets)
}

func (s *StreamXLSX) writeStylesheet() error {
	filename := "xl/styles.xml"
	fh, err := s.zip.Create(filename)
	if err != nil {
		return err
	}
	return writeStylesheet(fh, s.Styles)
}

func (s *StreamXLSX) writeSharedStrings() error {
	filename := "xl/sharedStrings.xml"
	fh, err := s.zip.Create(filename)
	if err != nil {
		return err
	}
	return writeSharedStrings(fh)
}

func (s *StreamXLSX) writeRelations() error {
	filename := "_rels/.rels"
	fh, err := s.zip.Create(filename)
	if err != nil {
		return err
	}
	return writeRelations(fh)
}

func (s *StreamXLSX) writeWorkbookRelations() error {
	filename := "xl/_rels/workbook.xml.rels"
	fh, err := s.zip.Create(filename)
	if err != nil {
		return err
	}
	return writeWorkbookRelations(fh, s.finishedSheets)
}

func (s *StreamXLSX) writeSheetRelations() error {
	if len(s.openSheet.relations) == 0 {
		return nil
	}
	filename := fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", len(s.finishedSheets)+1)
	fh, err := s.zip.Create(filename)
	if err != nil {
		return err
	}
	return writeRelations_(fh, s.openSheet.relations)
}

func (s *StreamXLSX) writeContentTypes() error {
	filename := "[Content_Types].xml"
	fh, err := s.zip.Create(filename)
	if err != nil {
		return err
	}
	// Pass the number of unique images for media overrides
	uniqueImageCount := 0
	if len(s.allImages) > 0 {
		// Count unique image names
		seen := make(map[string]bool)
		for _, img := range s.allImages {
			if !seen[img.name] {
				seen[img.name] = true
				uniqueImageCount++
			}
		}
	}
	return writeContentTypes(fh, len(s.finishedSheets), s.imageFormats, uniqueImageCount)
}

// writeImagesAndDrawings writes all image files and drawing XML files
// Returns a slice of image format strings for content types
func (s *StreamXLSX) writeImagesAndDrawings() ([]string, error) {
	// Collect all image formats for content types
	imageFormats := make(map[string]bool)

	// Map from image name to imageInfo
	imageMap := map[string]*imageInfo{}
	imageIndex := 0

	// First, load all unique images
	for _, imgRef := range s.allImages {
		if _, exists := imageMap[imgRef.name]; !exists {
			loadedImg, err := s.ImageLoader(imgRef.name)
			if err != nil {
				return nil, fmt.Errorf("failed to load image %q: %w", imgRef.name, err)
			}
			if loadedImg == nil {
				return nil, fmt.Errorf("ImageLoader returned nil for image %q", imgRef.name)
			}

			// Extract format from ContentType
			format := formatFromContentType(loadedImg.ContentType)
			if format == "" {
				format = "png" // default
			}
			imageFormats[format] = true

			// Use dimensions from LoadedImage
			width := loadedImg.Width
			height := loadedImg.Height
			// Override with Image struct dimensions if specified
			if imgRef.width > 0 {
				width = imgRef.width
			}
			if imgRef.height > 0 {
				height = imgRef.height
			}
			width = max(width, 20)
			height = max(height, 20)

			imageIndex++
			mediaPath := fmt.Sprintf("xl/media/image%d.%s", imageIndex, format)

			fh, err := s.zip.Create(mediaPath)
			if err != nil {
				return nil, fmt.Errorf("failed to create image file %s: %w", mediaPath, err)
			}
			if _, err = fh.Write(loadedImg.Data); err != nil {
				return nil, fmt.Errorf("failed to write image data for %s: %w", imgRef.name, err)
			}

			imageMap[imgRef.name] = &imageInfo{
				name:      imgRef.name,
				format:    format,
				mediaPath: mediaPath,
				width:     width,
				height:    height,
			}
		}
	}

	for sheetIdx, images := range s.sheetImages {
		if len(images) == 0 {
			continue
		}

		drawingPath := fmt.Sprintf("xl/drawings/drawing%d.xml", sheetIdx+1)
		fh, err := s.zip.Create(drawingPath)
		if err != nil {
			return nil, fmt.Errorf("failed to create drawing file %s: %w", drawingPath, err)
		}
		if err := s.writeDrawingXML(fh, images, imageMap); err != nil {
			return nil, err
		}

		relsPath := fmt.Sprintf("xl/drawings/_rels/drawing%d.xml.rels", sheetIdx+1)
		fh, err = s.zip.Create(relsPath)
		if err != nil {
			return nil, fmt.Errorf("failed to create drawing rels file %s: %w", relsPath, err)
		}

		drawingRels := make([]relationship, 0)
		for i, imgRef := range images {
			info := imageMap[imgRef.name]
			if info == nil {
				continue
			}
			// Relationship from drawing to image in media
			mediaRelPath := "../media/" + filepath.Base(info.mediaPath)
			drawingRels = append(drawingRels, relationship{
				ID:     fmt.Sprintf("rId%d", i+1),
				Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
				Target: mediaRelPath,
			})
		}

		if err := writeRelations_(fh, drawingRels); err != nil {
			return nil, fmt.Errorf("failed to write drawing relationships: %w", err)
		}
	}

	formatsList := make([]string, 0, len(imageFormats))
	for f := range imageFormats {
		formatsList = append(formatsList, f)
	}

	return formatsList, nil
}

// writeAllSheetRelations writes the .rels files for all sheets
func (s *StreamXLSX) writeAllSheetRelations() error {
	for sheetIdx := range s.sheetRelations {
		rels := s.sheetRelations[sheetIdx]

		if sheetIdx < len(s.sheetImages) && len(s.sheetImages[sheetIdx]) > 0 && s.ImageLoader != nil {
			rel := relationship{
				ID:     fmt.Sprintf("drawId%d", len(rels)+1),
				Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
				Target: fmt.Sprintf("/xl/drawings/drawing%d.xml", sheetIdx+1),
			}
			rels = append(rels, rel)
		}

		if len(rels) > 0 {
			filename := fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", sheetIdx+1)
			fh, err := s.zip.Create(filename)
			if err != nil {
				return err
			}
			if err := writeRelations_(fh, rels); err != nil {
				return err
			}
		}
	}
	return nil
}

// writeDrawingXML generates the drawing XML for a sheet's images
func (s *StreamXLSX) writeDrawingXML(w io.Writer, images []imageRef, imageMap map[string]*imageInfo) error {
	io.WriteString(w, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+"\n")
	io.WriteString(w, `<xdr:wsDr `+"\n")
	io.WriteString(w, `  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"`+"\n")
	io.WriteString(w, `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`+"\n")
	io.WriteString(w, `  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`+"\n")

	for i, imgRef := range images {
		info := imageMap[imgRef.name]
		if info == nil {
			continue
		}

		relID := fmt.Sprintf("rId%d", i+1)

		col := imgRef.anchorCol
		row := imgRef.anchorRow

		// Use dimensions from this imageRef, falling back to imageInfo
		// (which comes from LoadedImage, or the first Image struct seen)
		width := info.width
		height := info.height
		if imgRef.width > 0 {
			width = imgRef.width
		}
		if imgRef.height > 0 {
			height = imgRef.height
		}

		widthEMU := width * emuPerPixel
		heightEMU := height * emuPerPixel

		// Use oneCellAnchor like Google Sheets does
		io.WriteString(w, "  <xdr:oneCellAnchor>\n")
		// From point (top-left)
		io.WriteString(w, "    <xdr:from>\n")
		fmt.Fprintf(w, "      <xdr:col>%d</xdr:col>\n", col)
		io.WriteString(w, "      <xdr:colOff>0</xdr:colOff>\n")
		fmt.Fprintf(w, "      <xdr:row>%d</xdr:row>\n", row)
		io.WriteString(w, "      <xdr:rowOff>0</xdr:rowOff>\n")
		io.WriteString(w, "    </xdr:from>\n")
		// Size of the image
		fmt.Fprintf(w, "    <xdr:ext cx=\"%d\" cy=\"%d\"/>\n", widthEMU, heightEMU)
		// Picture
		io.WriteString(w, "    <xdr:pic>\n")
		io.WriteString(w, "      <xdr:nvPicPr>\n")
		fmt.Fprintf(w, "        <xdr:cNvPr id=\"%d\" name=\"%s\"/>\n", i+1, info.name)
		io.WriteString(w, "        <xdr:cNvPicPr preferRelativeResize=\"0\"/>\n")
		io.WriteString(w, "      </xdr:nvPicPr>\n")
		io.WriteString(w, "      <xdr:blipFill>\n")
		fmt.Fprintf(w, "        <a:blip r:embed=\"%s\"/>\n", relID)
		io.WriteString(w, "        <a:stretch>\n")
		io.WriteString(w, "          <a:fillRect/>\n")
		io.WriteString(w, "        </a:stretch>\n")
		io.WriteString(w, "      </xdr:blipFill>\n")
		io.WriteString(w, "      <xdr:spPr>\n")
		io.WriteString(w, "        <a:prstGeom prst=\"rect\">\n")
		io.WriteString(w, "          <a:avLst/>\n")
		io.WriteString(w, "        </a:prstGeom>\n")
		io.WriteString(w, "        <a:noFill/>\n")
		io.WriteString(w, "        <a:ln w=\"0\">\n")
		io.WriteString(w, "          <a:noFill/>\n")
		io.WriteString(w, "        </a:ln>\n")
		io.WriteString(w, "      </xdr:spPr>\n")
		io.WriteString(w, "    </xdr:pic>\n")
		io.WriteString(w, "    <xdr:clientData fLocksWithSheet=\"0\"/>\n")
		io.WriteString(w, "  </xdr:oneCellAnchor>\n")
	}

	io.WriteString(w, `</xdr:wsDr>`+"\n")
	return nil
}
