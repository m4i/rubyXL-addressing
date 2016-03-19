require 'codeclimate-test-reporter'
CodeClimate::TestReporter.start do
  add_filter '/test/'
end

$LOAD_PATH.unshift File.expand_path('../../lib', __FILE__)
require 'rubyXL/addressing'

require 'minitest/autorun'
