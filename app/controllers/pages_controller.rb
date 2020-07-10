class PagesController < ApplicationController
  def home
    @nbr_uniq_clients = uniq_instances_on_column(1)
    @nbr_uniq_orders = uniq_instances_on_column(10)
    @nbr_delivery_clik_and_collect = delivery_count("Click & Collect")
    @nbr_delivery_shopkeeper = delivery_count("Livraison par le commerçant")
    @nbr_delivery_la_poste = delivery_count("Livraison Laposte")
    @orders_per_client = orders_per_client
    @visits_per_month = visits_per_month
  end

  private

  def extract_data_from(file_path)
    Roo::Spreadsheet.open(file_path)
  end

  def uniq_instances_on_column(column_number)
    data = extract_data_from("lib/data/orders-test.xlsx")
    result = data.column(column_number).drop(1).uniq.count
  end

  def delivery_count(delivery_type)
    data = extract_data_from("lib/data/orders-test.xlsx")
    result = data.column(27).drop(1).count(delivery_type)
  end

  def orders_per_client
    data = extract_data_from("lib/data/orders-test.xlsx")
    result = {}
    data.each(order_id: 'Reference Beneficiaire(order_id)', raison_social: 'Raison Social') do |hash|
      next if hash[:order_id] == 'Reference Beneficiaire(order_id)'
      raison_social = hash[:raison_social]
      if result.has_key?(raison_social)
        result[raison_social] += 1
      else
        result[raison_social] = 1
      end 
    end
    result
  end

  def visits_per_month
    data = extract_data_from("lib/data/data-test.xlsx")
    result = {}
    data.each(date: 'Date', nbr_visits: 'Nb de visites') do |hash|
      next if hash[:date] == "Date"
      month = hash[:date].strftime("%b")
      if result.has_key?(month)
        result[month] += hash[:nbr_visits]
      else
        result[month] = hash[:nbr_visits]
      end 
    end
    result
  end
end 