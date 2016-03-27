# frozen_string_literal: true

require 'test_helper'
require 'rubyXL/address'

class RubyXL::AddressClassMethodsTest < Minitest::Test
  ROW_MAP = {
    0         => '1',
    1         => '2',
    1_048_574 => '1048575',
    1_048_575 => '1048576', # Excel max row
    1_048_576 => '1048577',
  }.freeze

  COLUMN_MAP = {
    0      => 'A',
    1      => 'B',
    25     => 'Z',
    26     => 'AA',
    16_382 => 'XFC',
    16_383 => 'XFD', # Excel max column
    16_384 => 'XFE',
  }.freeze

  def test_row_ind2ref
    ROW_MAP.each do |ind, ref|
      assert_equal ref, RubyXL::Address.row_ind2ref(ind)
    end
  end

  def test_column_ind2ref
    COLUMN_MAP.each do |ind, ref|
      assert_equal ref, RubyXL::Address.column_ind2ref(ind)
    end
  end

  def test_row_ref2ind
    ROW_MAP.each do |ind, ref|
      assert_equal ind, RubyXL::Address.row_ref2ind(ref)
      assert_equal ind, RubyXL::Address.row_ref2ind(ref.to_sym)
    end
  end

  def test_column_ref2ind
    COLUMN_MAP.each do |ind, ref|
      assert_equal ind, RubyXL::Address.column_ref2ind(ref)
      assert_equal ind, RubyXL::Address.column_ref2ind(ref.to_sym)
    end
  end

  def test_row_ind2ref_type_error
    assert_raises(TypeError) { RubyXL::Address.row_ind2ref(0.0) }
    assert_raises(TypeError) { RubyXL::Address.row_ind2ref(0r) }
    assert_raises(TypeError) { RubyXL::Address.row_ind2ref('0') }
  end

  def test_row_ind2ref_argument_error
    assert_raises(ArgumentError) { RubyXL::Address.row_ind2ref(-1) }
  end

  def test_column_ind2ref_type_error
    assert_raises(TypeError) { RubyXL::Address.column_ind2ref(0.0) }
    assert_raises(TypeError) { RubyXL::Address.column_ind2ref(0r) }
    assert_raises(TypeError) { RubyXL::Address.column_ind2ref('0') }
  end

  def test_column_ind2ref_argument_error
    assert_raises(ArgumentError) { RubyXL::Address.column_ind2ref(-1) }
  end

  def test_row_ref2ind_type_error
    assert_raises(TypeError) { RubyXL::Address.row_ref2ind(1) }
  end

  def test_row_ref2ind_argument_error
    assert_raises(ArgumentError) { RubyXL::Address.row_ref2ind('0') }
    assert_raises(ArgumentError) { RubyXL::Address.row_ref2ind('01') }
  end

  def test_column_ref2ind_type_error
    assert_raises(TypeError) { RubyXL::Address.column_ref2ind(1) }
  end

  def test_column_ref2ind_argument_error
    assert_raises(ArgumentError) { RubyXL::Address.column_ref2ind('1') }
    assert_raises(ArgumentError) { RubyXL::Address.column_ref2ind('a') }
  end
end
