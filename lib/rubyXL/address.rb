# frozen_string_literal: true

module RubyXL
  class Address
    ROW_REF_FORMAT    = /\A[1-9]\d*\z/
    COLUMN_REF_FORMAT = /\A[A-Z]+\z/

    class << self
      # @param [Integer] ind
      # @return [String]
      def row_ind2ref(ind)
        message = "invalid row #{ind.inspect}"
        raise TypeError,     message unless ind.is_a?(Integer)
        raise ArgumentError, message unless ind >= 0

        (ind + 1).to_s.freeze
      end

      # @param [Integer] ind
      # @return [String]
      def column_ind2ref(ind)
        message = "invalid column #{ind.inspect}"
        raise TypeError,     message unless ind.is_a?(Integer)
        raise ArgumentError, message unless ind >= 0

        ref = ''.dup
        loop do
          ref.prepend((ind % 26 + 65).chr)
          ind /= 26
          break if (ind -= 1) < 0
        end
        ref.freeze
      end

      # @param [String, Symbol] ref
      # @return [Integer]
      def row_ref2ind(ref)
        message = "invalid row #{ref.inspect}"
        raise ArgumentError, message unless ROW_REF_FORMAT =~ ref

        ref.to_s.to_i - 1
      end

      # @param [String, Symbol] ref
      # @return [Integer]
      def column_ref2ind(ref)
        message = "invalid column #{ref.inspect}"
        raise ArgumentError, message unless COLUMN_REF_FORMAT =~ ref

        ref.to_s.each_byte.reduce(0) { |a, e| a * 26 + (e - 64) } - 1
      end
    end

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
      case row
      when Integer
        return @row = row if row >= 0
      when String, Symbol
        return self.row = self.class.row_ref2ind(row)
      end

      raise ArgumentError, "invalid row #{row.inspect}"
    end

    # @param [Integer, String, Symbol] column
    # @return [Integer, String, Symbol]
    def column=(column)
      case column
      when Integer
        return @column = column if column >= 0
      when String, Symbol
        return self.column = self.class.column_ref2ind(column)
      end

      raise ArgumentError, "invalid column #{column.inspect}"
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

    # @return [Boolean]
    def exists?
      !cell.nil?
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
