# rubyXL-addressing

Addressing for [rubyXL](https://github.com/weshatheleopard/rubyXL)

[![Build Status](https://travis-ci.org/m4i/rubyXL-addressing.svg?branch=master)](https://travis-ci.org/m4i/rubyXL-addressing)
[![Test Coverage](https://codeclimate.com/github/m4i/rubyXL-addressing/badges/coverage.svg)](https://codeclimate.com/github/m4i/rubyXL-addressing/coverage)
[![Dependency Status](https://gemnasium.com/m4i/rubyXL-addressing.svg)](https://gemnasium.com/m4i/rubyXL-addressing)

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'rubyXL-addressing'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install rubyXL-addressing

## Usage

### calculate C3 + F3

#### rubyXL

```ruby
row_index, column_index_c = RubyXL::Reference.ref2ind('C3')
_,         column_index_f = RubyXL::Reference.ref2ind('F3')
row = worksheet[row_index]
if row
  cell_c = row[column_index_c]
  cell_f = row[column_index_f]
  if cell_c && cell_f
    cell_c.value + cell_f.value
  end
end
```

#### rubyXL-addressing

```ruby
c3 = worksheet.addr(:C3)
f3 = c3.column(:F)          # or c3.right(3)
if c3.cell && f3.cell       # or c3.exists?
  c3.value + f3.value
end
```

### set value to C3

#### rubyXL

```ruby
row_index, column_index = RubyXL::Reference.ref2ind('C3')
row = worksheet[row_index]
if row && row[column_index]
  row[column_index].change_contents('foobar')
else
  worksheet.add_cell(row_index, column_index, 'foobar')
end
```

#### rubyXL-addressing

```ruby
worksheet.addr('C3').value = 'foobar'
```

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake test` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/rubyXL-addressing. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [Contributor Covenant](http://contributor-covenant.org) code of conduct.

## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).
