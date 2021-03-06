# frozen_string_literal: true

require 'test_helper'
require 'rubyXL/address'

class RubyXL::AddressTest < Minitest::Test
  def test_worksheet
    worksheet1 = 'dummy1'
    worksheet2 = 'dummy2'

    addr = RubyXL::Address.new(worksheet1, ref: :A1)
    assert_same worksheet1, addr.worksheet

    # #worksheet(value) creates an new Address.
    assert_same worksheet2, addr.worksheet(worksheet2).worksheet

    # #worksheet(value) does not change self.
    assert_same worksheet1, addr.worksheet
  end

  def test_worksheet_setter
    worksheet1 = 'dummy1'
    worksheet2 = 'dummy2'

    addr = RubyXL::Address.new(worksheet1, ref: :A1)

    addr.worksheet = worksheet2
    assert_same worksheet2, addr.worksheet
  end

  def test_row
    addr = RubyXL::Address.new('dummy', row: 2, column: 1)
    assert_equal 2, addr.row

    # #row(value) creates an new Address.
    assert_equal 5, addr.row(5).row
    assert_equal 4, addr.row('5').row
    assert_equal 6, addr.row(:'7').row

    # #row(value) does not change column.
    assert_equal addr.column, addr.row(5).column

    # #row(value) does not change self.
    assert_equal 2, addr.row
  end

  def test_column
    addr = RubyXL::Address.new('dummy', row: 1, column: 2)
    assert_equal 2, addr.column

    # #column(value) creates an new Address.
    assert_equal 5, addr.column(5).column
    assert_equal 4, addr.column('E').column
    assert_equal 6, addr.column(:G).column

    # #column(value) does not change row.
    assert_equal addr.row, addr.column(5).row

    # #column(value) does not change self.
    assert_equal 2, addr.column
  end

  def test_row_setter
    addr = RubyXL::Address.new('dummy', ref: :A1)

    addr.row = 5
    assert_equal 5, addr.row

    addr.row = '5'
    assert_equal 4, addr.row

    addr.row = :'7'
    assert_equal 6, addr.row
  end

  def test_column_setter
    addr = RubyXL::Address.new('dummy', ref: :A1)

    addr.column = 5
    assert_equal 5, addr.column

    addr.column = 'E'
    assert_equal 4, addr.column

    addr.column = :G
    assert_equal 6, addr.column
  end

  def test_ref
    assert_equal 'A1', RubyXL::Address.new('dummy', ref: :A1).ref
  end

  def test_inspect
    worksheet = workbook.add_worksheet('TEST')
    assert_equal '#<RubyXL::Address TEST!A1>',
                 RubyXL::Address.new(worksheet, ref: :A1).inspect
  end

  def test_up
    addr = RubyXL::Address.new('dummy', ref: :A9)
    assert_equal addr.row - 1, addr.up.row
    assert_equal addr.row - 0, addr.up(0).row
    assert_equal addr.row - 1, addr.up(1).row
    assert_equal addr.row - 3, addr.up(3).row
    assert_equal addr.row + 3, addr.up(-3).row
    assert_equal addr.column, addr.up.column
  end

  def test_up_error
    addr = RubyXL::Address.new('dummy', row: 0, column: 1)
    assert_raises(ArgumentError) { addr.up }
  end

  def test_down
    addr = RubyXL::Address.new('dummy', ref: :A9)
    assert_equal addr.row + 1, addr.down.row
    assert_equal addr.row + 0, addr.down(0).row
    assert_equal addr.row + 1, addr.down(1).row
    assert_equal addr.row + 3, addr.down(3).row
    assert_equal addr.row - 3, addr.down(-3).row
    assert_equal addr.column, addr.down.column
  end

  def test_left
    addr = RubyXL::Address.new('dummy', ref: :I1)
    assert_equal addr.column - 1, addr.left.column
    assert_equal addr.column - 0, addr.left(0).column
    assert_equal addr.column - 1, addr.left(1).column
    assert_equal addr.column - 3, addr.left(3).column
    assert_equal addr.column + 3, addr.left(-3).column
    assert_equal addr.row, addr.left.row
  end

  def test_left_error
    addr = RubyXL::Address.new('dummy', row: 1, column: 0)
    assert_raises(ArgumentError) { addr.left }
  end

  def test_right
    addr = RubyXL::Address.new('dummy', ref: :I1)
    assert_equal addr.column + 1, addr.right.column
    assert_equal addr.column + 0, addr.right(0).column
    assert_equal addr.column + 1, addr.right(1).column
    assert_equal addr.column + 3, addr.right(3).column
    assert_equal addr.column - 3, addr.right(-3).column
    assert_equal addr.row, addr.right.row
  end

  def test_up!
    addr = RubyXL::Address.new('dummy', ref: :A9)
    row    = addr.row
    column = addr.column

    addr.up!(1)
    row -= 1
    assert_equal row, addr.row

    addr.up!(0)
    assert_equal row, addr.row

    addr.up!(1)
    row -= 1
    assert_equal row, addr.row

    addr.up!(3)
    row -= 3
    assert_equal row, addr.row

    addr.up!(-3)
    row += 3
    assert_equal row, addr.row

    # #up! does not change column.
    assert_equal column, addr.column
  end

  def test_up_error!
    addr = RubyXL::Address.new('dummy', row: 0, column: 1)
    assert_raises(ArgumentError) { addr.up! }
  end

  def test_down!
    addr = RubyXL::Address.new('dummy', ref: :A9)
    row    = addr.row
    column = addr.column

    addr.down!(1)
    row += 1
    assert_equal row, addr.row

    addr.down!(0)
    assert_equal row, addr.row

    addr.down!(1)
    row += 1
    assert_equal row, addr.row

    addr.down!(3)
    row += 3
    assert_equal row, addr.row

    addr.down!(-3)
    row -= 3
    assert_equal row, addr.row

    # #down! does not change column.
    assert_equal column, addr.column
  end

  def test_left!
    addr = RubyXL::Address.new('dummy', ref: :I1)
    row    = addr.row
    column = addr.column

    addr.left!(1)
    column -= 1
    assert_equal column, addr.column

    addr.left!(0)
    assert_equal column, addr.column

    addr.left!(1)
    column -= 1
    assert_equal column, addr.column

    addr.left!(3)
    column -= 3
    assert_equal column, addr.column

    addr.left!(-3)
    column += 3
    assert_equal column, addr.column

    # #left! does not change row.
    assert_equal row, addr.row
  end

  def test_left_error!
    addr = RubyXL::Address.new('dummy', row: 1, column: 0)
    assert_raises(ArgumentError) { addr.left! }
  end

  def test_right!
    addr = RubyXL::Address.new('dummy', ref: :I1)
    row    = addr.row
    column = addr.column

    addr.right!(1)
    column += 1
    assert_equal column, addr.column

    addr.right!(0)
    assert_equal column, addr.column

    addr.right!(1)
    column += 1
    assert_equal column, addr.column

    addr.right!(3)
    column += 3
    assert_equal column, addr.column

    addr.right!(-3)
    column -= 3
    assert_equal column, addr.column

    # #right! does not change row.
    assert_equal row, addr.row
  end

  def test_cell
    addr = RubyXL::Address.new(worksheet, ref: :A9)
    assert_nil addr.cell
    worksheet.add_cell(addr.row, addr.column)
    assert_same worksheet[addr.row][addr.column], addr.cell
  end

  def test_value
    addr = RubyXL::Address.new(worksheet, ref: :C7)
    assert_nil addr.value
    worksheet.add_cell(addr.row, addr.column, 'foobar')
    assert_same worksheet[addr.row][addr.column].value, addr.value
  end

  def test_value_setter
    addr = RubyXL::Address.new(worksheet, ref: :D6)
    assert_nil addr.cell

    value = 'foo'
    addr.value = value
    assert_same value, worksheet[addr.row][addr.column].value

    value = 'baz'
    addr.value = value
    assert_same value, worksheet[addr.row][addr.column].value

    addr = RubyXL::Address.new(worksheet, ref: :E5)
    assert_nil addr.cell

    value = true
    addr.value = value
    assert_same value, worksheet[addr.row][addr.column].value
  end

  private

  def workbook
    @@workbook ||=
      begin
        require 'rubyXL'
        RubyXL::Workbook.new
      end
  end

  def worksheet
    @@worksheet ||= workbook[0]
  end
end
