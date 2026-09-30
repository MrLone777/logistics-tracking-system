CREATE TABLE IF NOT EXISTS shipments (
  id BIGSERIAL PRIMARY KEY,
  category VARCHAR(20) NOT NULL CHECK (category IN ('Local','Domestic','International')),
  shipment_type VARCHAR(30),
  name VARCHAR(200),
  crn_no VARCHAR(100),
  eway_bill_no VARCHAR(100),
  delivery_challan_no VARCHAR(100),
  tracking_lr_no VARCHAR(100),
  tracking_id VARCHAR(150),
  gst VARCHAR(50),
  amount NUMERIC(14,2),
  origin VARCHAR(200),
  destination VARCHAR(200),
  status VARCHAR(40) NOT NULL DEFAULT 'Pending',
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_shipments_category ON shipments(category);
CREATE INDEX IF NOT EXISTS idx_shipments_tracking_id ON shipments(tracking_id);
CREATE INDEX IF NOT EXISTS idx_shipments_crn_no ON shipments(crn_no);
CREATE INDEX IF NOT EXISTS idx_shipments_eway_bill_no ON shipments(eway_bill_no);
