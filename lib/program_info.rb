class ProgInfo
  
  def initialize(rating, year, data)
    @rating = rating
    @year = year
    @data = data
  end

  def get_match(pattern)
    match = ""
    @data.split.each do |d|
      if pattern.match(d)
        match = cleanup_regex(d)
      end
    end
    return match
  end

  def rating
    pattern = Regexp.new(@rating)
    match = get_match(pattern)
  end

  def year
    pattern = Regexp.new(@year)
    match = get_match(pattern)
  end

  def cleanup_regex(data)
    string = data.sub(/^(\[|\()(.*)(\]|\))$/, '\2')
  end
end