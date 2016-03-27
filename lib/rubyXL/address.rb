# frozen_string_literal: true

require 'rubyXL/objects/reference'

module RubyXL
  class Address
    ROW_REF_FORMAT    = /\A[1-9]\d*\z/
    COLUMN_REF_FORMAT = /\A[A-Z]+\z/

    class << self
      # @param [Integer] ind
      # @return [String]
      def row_ind2ref(ind)
        validate_index(:row, ind)

        (ind + 1).to_s.freeze
      end

      # @param [Integer] ind
      # @return [String]
      def column_ind2ref(ind)
        validate_index(:column, ind)

        ref = ''.dup
        loop do
          ref.prepend((ind % 26 + 65).chr)
          ind = ind / 26 - 1
          break if ind < 0
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

      private

      def normalize(row_or_column, value)
        case value
        when String, Symbol
          value = public_send("#{row_or_column}_ref2ind", value)
        else
          validate_index(row_or_column, value)
        end
        value
      end

      def validate_index(row_or_column, index)
        message = "invalid #{row_or_column} #{index.inspect}"
        raise TypeError,     message unless index.is_a?(Integer)
        raise ArgumentError, message unless index >= 0
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

      row, column = RubyXL::Reference.ref2ind(ref) if ref
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
      @row = self.class.__send__(:normalize, :row, row)
    end

    # @param [Integer, String, Symbol] column
    # @return [Integer, String, Symbol]
    def column=(column)
      @column = self.class.__send__(:normalize, :column, column)
    end

    # @return [String]
    def ref
      RubyXL::Reference.ind2ref(@row, @column)
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
