package streamxlsx

import (
	"strings"
)

// ImageLoader is called during Close() to fetch image data.
// The name passed is the Name field from the Image struct.
type (
	ImageLoader func(name string) (*LoadedImage, error)
	LoadedImage struct {
		Data        []byte
		ContentType string
		Width       int
		Height      int
	}

	// Image represents an image to be embedded in the worksheet.
	// The image is anchored to the cell where it appears in WriteRow().
	Image struct {
		Name   string // Identifier for the image, passed to ImageLoader
		Width  int    // Render width in pixels; if Width and Height are 0, scales to (kinda) fit a cell
		Height int    // Render height in pixels
	}
)

// imageRef is an internal struct for tracking images during writing
type imageRef struct {
	name       string
	anchorCol  int // 0-based column index
	anchorRow  int // 0-based row index
	sheetIndex int // which sheet this image belongs to
	width      int // image width in pixels (if known)
	height     int // image height in pixels (if known)
}

// imageInfo holds image metadata (dimensions, format, path)
type imageInfo struct {
	name      string
	format    string // png, jpeg, gif
	mediaPath string // e.g., "xl/media/image1.png"
	width     int    // image width in pixels
	height    int    // image height in pixels
}

// emuPerPixel is the conversion factor (1 pixel = 9525 EMUs)
// Excel uses: pixels * 9525 for EMU dimensions
const emuPerPixel = 9525

// formatFromContentType extracts the image format from a MIME content type.
// e.g., "image/png" -> "png", "image/jpeg" -> "jpeg"
func formatFromContentType(contentType string) string {
	// Remove "image/" prefix
	format := strings.TrimPrefix(contentType, "image/")
	if format == "jpg" {
		return "jpeg"
	}
	return format
}
