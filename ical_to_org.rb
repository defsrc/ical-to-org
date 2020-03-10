# frozen_string_literal: true

require 'bundler/inline'
require 'date'
require 'delegate'
require 'erb'
require 'open-uri'

gemfile do
  source 'https://rubygems.org'
  gem 'activesupport', '~>6.0'
  gem 'awesome_print', '~>1.8'
  gem 'icalendar', '~>2.6'
  gem 'icalendar-recurrence', '~>1.1'
end

URL = ENV.fetch('OUTLOOK_TO_ORG_URL')
OUTPUT_FILE = ENV.fetch('OUTLOOK_TO_ORG_OUTPUT')
MIN_DATE = 1.week.ago
MAX_DATE = 10.days.from_now
TIMEZONE = 'Europe/Berlin'

# If _one_ of the rejection reasons returns true
# the single event is rejected and will not be part
# of the output.
REJECTION_REASONS = [
  # reject every event that is not confirmed
  ->(e) { e.status != 'CONFIRMED' },
  # reject all all-day events
  ->(e) { e.all_day? },
  # reject every event that is titled like a blocker
  ->(e) { e.summary == '[GTD Blocker]' }
].freeze

class OrgEvent < SimpleDelegator
  attr_reader :start_time, :end_time

  def initialize(event, start_time, end_time)
    @start_time, @end_time = start_time, end_time
    super(event)
  end

  def uniq_key
    "#{duration_timestamp}::#{summary}"
  end

  def all_day?
    dtstart.is_a?(Icalendar::Values::Date)
  end

  def duration_timestamp
    format_time = ->(dt) { dt.in_time_zone(TIMEZONE).strftime('<%Y-%m-%d %a %H:%M>') }
    [start_time, end_time].map(&format_time).join('--')
  end
end

class OrgCalendar
  attr_reader :events

  def save
    content = ERB.new(DATA.read).result(binding)
    File.write(OUTPUT_FILE, content)
  end

  def load_events
    @events = open(URL).read
                       .then(&method(:repair_time_zones))
                       .then(&method(:parse_ical_events))
                       .flat_map(&method(:generate_org_events_per_occurrence))
                       .uniq(&:uniq_key)
                       .reject(&method(:reject_by_filter?))
  end

  private

  def reject_by_filter?(event)
    REJECTION_REASONS.any? { |check| check.call(event) }
  end

  def repair_time_zones(response)
    response
      .gsub('TZID=W. Europe Standard Time', 'TZID=Europe/Berlin')
      .gsub('TZID=GMT Standard Time', 'TZID=Europe/London')
  end

  def generate_org_events_per_occurrence(ical_event)
    ical_event.occurrences_between(MIN_DATE, MAX_DATE).map do |occurrence|
      OrgEvent.new(ical_event, occurrence.start_time, occurrence.end_time)
    end
  end

  def parse_ical_events(response)
    Icalendar::Calendar.parse(response).flat_map(&:events)
  end
end

calendar = OrgCalendar.new.tap do |c|
  c.load_events
  c.save
end

puts "Success! #{calendar.events.count} events were written to #{OUTPUT_FILE}"

__END__
# -*- buffer-read-only: t -*-
#+TITLE: Outlook Calendar Export
#+CATEGORY: Cal
#+SETUPFILE: /Users/fabrik42/org/_ioki_config.org
#+ICAL_EXPORT_DATE: <%= Time.now.iso8601 %>

<% events.each do |event| %>
* <%= event.summary %>
:PROPERTIES:
:ICAL_UID: <%= event.uid %>
:ICAL_LOCATION: <%= event.location %>
:ICAL_START: <%= event.start_time %>
:ICAL_END: <%= event.end_time %>
:ICAL_DTSTAMP: <%= event.dtstamp %>
:END:
<%= event.duration_timestamp %>
Location: <%= event.location %>
<% end %>
