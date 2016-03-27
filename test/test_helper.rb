# frozen_string_literal: true

require 'codeclimate-test-reporter'
CodeClimate::TestReporter.start do
  add_filter '/test/'
end

$LOAD_PATH.unshift File.expand_path('../../lib', __FILE__)

require 'minitest/autorun'
