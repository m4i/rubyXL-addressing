# frozen_string_literal: true

require 'rubyXL/addressing/version'

module RubyXL
  autoload :Address, 'rubyXL/address'

  module Addressing
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

    # @param [String, Symbol, Integer] ref_or_row is ref `'A1'`
    #                                             or row index `0`
    #                                             or row label `'1'`
    # @param [String, Symbol, Integer] column is column index `0`
    #                                         or column label `'A'`
    # @return [RubyXL::Address]
    def addr(ref_or_row, column = nil)
      if column.nil?
        RubyXL::Address.new(self, ref: ref_or_row)
      else
        RubyXL::Address.new(self, row: ref_or_row, column: column)
      end
    end
  end
end

require 'rubyXL'
RubyXL::Worksheet.include(RubyXL::Addressing)
