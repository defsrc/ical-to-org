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
OUTPUT_FILE = '/Users/fabrik42/org/Cal.org'
MIN_DATE = 1.week.ago
MAX_DATE = 10.days.from_now
TIMEZONE = 'Europe/Berlin'

# If _one_ of the rejection filters return true
# the single event is rejected and will not be part
# of the output.
REJECTION_FILTERS = [
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

  def confirmed?
    status == 'CONFIRMED'
  end

  def rejected_by_filter?
    REJECTION_FILTERS.any? { |check| check.call(self) }
  end

  def duration_timestamp
    format_time = ->(dt) { dt.in_time_zone(TIMEZONE).strftime('<%Y-%m-%d %a %H:%M>') }
    [start_time, end_time].map(&format_time).join('--')
  end

  def render
    ERB.new(template).result(binding)
  end

  def template
    <<~ERB
      * <%= summary %>
      :PROPERTIES:
      :ICAL_UID: <%= uid %>
      :ICAL_LOCATION: <%= location %>
      :ICAL_START: <%= start_time %>
      :ICAL_END: <%= end_time %>
      :ICAL_DTSTAMP: <%= dtstamp %>
      :END:
      <%= duration_timestamp %>
      Location: <%= location %>
    ERB
  end
end

def repair_time_zones(response)
  response.gsub('TZID=W. Europe Standard Time', 'TZID=Europe/Berlin')
end

def generate_org_events_per_occurrence(ical_event)
  ical_event.occurrences_between(MIN_DATE, MAX_DATE).map do |occurrence|
    OrgEvent.new(ical_event, occurrence.start_time, occurrence.end_time)
  end
end

def parse_ical_events(response)
  Icalendar::Calendar.parse(response).flat_map(&:events)
end

events = open(URL).read
                  .then(&method(:repair_time_zones))
                  .then(&method(:parse_ical_events))
                  .flat_map(&method(:generate_org_events_per_occurrence))
                  .uniq(&:uniq_key)
                  .select(&:confirmed?)
                  .reject(&:all_day?)
                  .reject(&:rejected_by_filter?)

File.open(OUTPUT_FILE, 'w') do |f|
  f << <<~HEADER
    # -*- buffer-read-only: t -*-
    #+TITLE: Outlook Calendar Export
    #+CATEGORY: Cal
    #+ICAL_EXPORT_DATE: #{Time.now.iso8601}
    #+SETUPFILE: /Users/fabrik42/org/_ioki_config.org

  HEADER

  events.each { |e| f << e.render }
end

puts "Success! #{events.count} events were written to #{OUTPUT_FILE}"
