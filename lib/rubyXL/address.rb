# frozen_string_literal: true

module RubyXL
  class Address
    # @param [RubyXL::Worksheet] worksheet
    # @return [RubyXL::Worksheet]
    attr_writer :worksheet

    # @param [RubyXL::Worksheet] worksheet
    # @param [String, Symbol] ref
    # @param [Integer, String, Symbol] row
    # @param [Integer, String, Symbol] column
    def initialize(worksheet, ref: nil, row: nil, column: nil)
      @worksheet = worksheet

      row, column = Reference.ref2ind(ref) if ref
      self.row    = row
      self.column = column
    end

    # @param [RubyXL::Worksheet, nil] worksheet
    # @return [RubyXL::Workbook, RubyXL::Address]
    def worksheet(worksheet = nil)
      if worksheet.nil?
        @worksheet
      else
        self.class.new(worksheet, row: @row, column: @column)
      end
    end

    # @param [Integer, String, Symbol, nil] row
    # @return [Integer, RubyXL::Address]
    def row(row = nil)
      if row.nil?
        @row
      else
        self.class.new(@worksheet, row: row, column: @column)
      end
    end

    # @param [Integer, String, Symbol, nil] column
    # @return [Integer, RubyXL::Address]
    def column(column = nil)
      if column.nil?
        @column
      else
        self.class.new(@worksheet, row: @row, column: column)
      end
    end

    # @param [Integer, String, Symbol] row
    # @return [Integer, String, Symbol]
    def row=(row)
      @row = Addressing.__send__(:normalize, :row, row)
    end

    # @param [Integer, String, Symbol] column
    # @return [Integer, String, Symbol]
    def column=(column)
      @column = Addressing.__send__(:normalize, :column, column)
    end

    # @return [String]
    def ref
      Reference.ind2ref(@row, @column)
    end

    # @return [String]
    def inspect
      format('#<%s %s!%s>', self.class.name, @worksheet.sheet_name, ref)
    end

    # @param [Integer] amount
    # @return [RubyXL::Address]
    def up(amount = 1)
      row(@row - amount)
    end

    # @param [Integer] amount
    # @return [RubyXL::Address]
    def down(amount = 1)
      row(@row + amount)
    end

    # @param [Integer] amount
    # @return [RubyXL::Address]
    def left(amount = 1)
      column(@column - amount)
    end

    # @param [Integer] amount
    # @return [RubyXL::Address]
    def right(amount = 1)
      column(@column + amount)
    end

    # @param [Integer] amount
    # @return [self]
    def up!(amount = 1)
      self.row -= amount
      self
    end

    # @param [Integer] amount
    # @return [self]
    def down!(amount = 1)
      self.row += amount
      self
    end

    # @param [Integer] amount
    # @return [self]
    def left!(amount = 1)
      self.column -= amount
      self
    end

    # @param [Integer] amount
    # @return [self]
    def right!(amount = 1)
      self.column += amount
      self
    end

    # @return [RubyXL::Cell, nil]
    def cell
      (row = @worksheet[@row]) && row[@column]
    end

    # @return [Object]
    def value
      cell && cell.value
    end

    # @param [Object] value
    # @return [Object]
    def value=(value)
      if cell
        cell.change_contents(value)
      else
        @worksheet.add_cell(@row, @column, value)
      end
    end
  end
end
