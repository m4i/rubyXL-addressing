# frozen_string_literal: true

require 'rubyXL/addressing/version'

module RubyXL
  autoload :Address, 'rubyXL/address'

  module Addressing
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
