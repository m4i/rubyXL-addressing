# frozen_string_literal: true

require 'test_helper'

class RubyXL::AddressingTest < Minitest::Test
  include RubyXL::Addressing

  def test_that_it_has_a_version_number
    refute_nil ::RubyXL::Addressing::VERSION
  end

  def test_addr
    addr = addr(:B2)
    assert_instance_of RubyXL::Address, addr
    assert_equal 'B2', addr.ref

    addr = addr(1, 2)
    assert_instance_of RubyXL::Address, addr
    assert_equal 1, addr.row
    assert_equal 2, addr.column
  end
end
