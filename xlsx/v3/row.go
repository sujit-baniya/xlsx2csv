package xlsx

import (
	"fmt"
)

// Row represents a single Row in the current Sheet.
type Row struct {
	Hidden       bool    // Hidden determines whether this Row is hidden or not.
	Sheet        *Sheet  // Sheet is a reference back to the Sheet that this Row is within.
	height       float64 // Height is the current height of the Row in PostScript Points
	outlineLevel uint8   // OutlineLevel contains the outline level of this Row.  Used for collapsing.
	isCustom     bool    // isCustom is a flag that is set to true when the Row has been modified
	num          int     // Num hold the positional number of the Row in the Sheet
	cellCount    int     // The current number of cells
	cells        []*Cell // the cells
}

// SetHeight sets the height of the Row in PostScript points
func (r *Row) SetHeight(ht float64) {
	r.height = ht
	r.isCustom = true
}

func (r *Row) GetCellCount() int {
    return r.cellCount;
}

// SetHeightCM sets the height of the Row in centimetres, inherently converting it to PostScript points.
func (r *Row) SetHeightCM(ht float64) {
	r.height = ht * 28.3464567 // Convert CM to postscript points
	r.isCustom = true
}

// GetHeight returns the height of the Row in PostScript points.
func (r *Row) GetHeight() float64 {
	return r.height
}

// SetOutlineLevel sets the outline level of the Row (used for collapsing rows)
func (r *Row) SetOutlineLevel(outlineLevel uint8) {
	r.outlineLevel = outlineLevel
	if r.Sheet != nil {
		if r.outlineLevel > r.Sheet.SheetFormat.OutlineLevelRow {
			r.Sheet.SheetFormat.OutlineLevelRow = outlineLevel
		}
	}
}

// GetOutlineLevel returns the outline level of the Row.
func (r *Row) GetOutlineLevel() uint8 {
	return r.outlineLevel
}

// AddCell adds a new Cell to the Row
func (r *Row) AddCell() *Cell {
	cell := newCell(r, r.cellCount)
	r.cellCount++
	r.cells = append(r.cells, cell)
	return cell
}

func (r *Row) makeCellKey(colIdx int) string {
	return fmt.Sprintf("%s:%06d:%06d", r.Sheet.Name, r.num, colIdx)
}

func (r *Row) key() string {
	return r.makeCellKeyRowPrefix()
}

func (r *Row) makeCellKeyRowPrefix() string {
	return fmt.Sprintf("%s:%06d", r.Sheet.Name, r.num)
}

func (r *Row) growCellsSlice(newSize int) {
	capacity := cap(r.cells)
	if newSize >= capacity {
		newCap := 2 * capacity
		if newSize > newCap {
			newCap = newSize
		}
		newSlice := make([]*Cell, newCap, newCap)
		copy(newSlice, r.cells)
		r.cells = newSlice
	}
}

// GetCell returns the Cell at a given column index, creating it if it doesn't exist.
func (r *Row) GetCell(colIdx int) *Cell {
	if colIdx >= len(r.cells) {
		cell := newCell(r, colIdx)
		r.growCellsSlice(colIdx + 1)

		r.cells[colIdx] = cell
		return cell
	}

	cell := r.cells[colIdx]
	if cell == nil {
		cell = newCell(r, colIdx)
		r.cells[colIdx] = cell
	}
	return cell
}

// cellVisitorFlags contains flags that can be set by CellVisitorOption implementations to modify the behaviour of ForEachCell
type cellVisitorFlags struct {
	// skipEmptyCells indicates if we should skip nil cells.
	skipEmptyCells bool
}

// CellVisitorOption describes a function that can set values in a
// cellVisitorFlags struct to affect the way ForEachCell operates
type CellVisitorOption func(flags *cellVisitorFlags)

// SkipEmptyCells can be passed as an option to Row.ForEachCell in
// order to make it skip over empty cells in the sheet.
func SkipEmptyCells(flags *cellVisitorFlags) {
	flags.skipEmptyCells = true
}

// ForEachCell will call the provided CellVisitorFunc for each
// currently defined cell in the Row.  Optionally you may pass one or
// more CellVisitorOption to affect how ForEachCell operates.  For
// example you may wish to pass SkipEmptyCells to only visit cells
// which are populated.
func (r *Row) ForEachCell(cvf CellVisitorFunc, option ...CellVisitorOption) error {
	flags := &cellVisitorFlags{}
	for _, opt := range option {
		opt(flags)
	}
	fn := func(ci int, c *Cell) error {
		if c == nil {
			if flags.skipEmptyCells {
				return nil
			}
			c = r.GetCell(ci)
		}
		if c.Value == "" && flags.skipEmptyCells {
			return nil
		}
		c.Row = r
		return cvf(c)
	}

	for ci, cell := range r.cells {
		err := fn(ci, cell)
		if err != nil {
			return err
		}
	}
	cellCount := len(r.cells)
	if !flags.skipEmptyCells {
		for ci := cellCount; ci < r.Sheet.MaxCol; ci++ {
			c := r.GetCell(ci)
			err := cvf(c)
			if err != nil {
				return err
			}

		}
	}

	return nil
}
