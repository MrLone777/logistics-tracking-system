import { useEffect, useState } from "react";
import * as XLSX from "xlsx";

const API_BASE = import.meta.env.VITE_API_BASE_URL || "http://localhost:4000";

function App() {
  const [apiError, setApiError] = useState("");
  const [isLoadingShipments, setIsLoadingShipments] = useState(true);
  const [activeCategory, setActiveCategory] = useState("All");
  const [showAddForm, setShowAddForm] = useState(false);
  const [showImportForm, setShowImportForm] = useState(false);
  const [importResult, setImportResult] = useState(null);
  const [isImporting, setIsImporting] = useState(false);
  const [selectedTrackingShipment, setSelectedTrackingShipment] = useState(null);

  const [shipments, setShipments] = useState([
    {
      id: 1,
      category: "Local",
      subType: "Porter",
      carrier: "Porter",
      crnNo: "PORT001",
      ewayBillNo: "",
      deliveryChallanNo: "",
      trackingNo: "",
      origin: "Chennai",
      destination: "Chennai",
      status: "In Transit",
      updated: "30 Sep 2026",
    },
    {
      id: 2,
      category: "Domestic",
      subType: "Carrier / Transporter",
      carrier: "FedEx",
      crnNo: "",
      ewayBillNo: "EWB123456789",
      deliveryChallanNo: "DC12345",
      trackingNo: "FDX123456",
      origin: "Chennai",
      destination: "Mumbai",
      status: "In Transit",
      updated: "30 Sep 2026",
    },
    {
      id: 3,
      category: "Domestic",
      subType: "Carrier / Transporter",
      carrier: "BlueDart",
      crnNo: "",
      ewayBillNo: "EWB987654321",
      deliveryChallanNo: "DC78945",
      trackingNo: "",
      origin: "Bangalore",
      destination: "Chennai",
      status: "Pending",
      updated: "30 Sep 2026",
    },
    {
      id: 4,
      category: "International",
      subType: "International Shipment",
      carrier: "Maersk",
      crnNo: "",
      ewayBillNo: "",
      deliveryChallanNo: "",
      trackingNo: "MSKU1234567",
      origin: "Qingdao, China",
      destination: "Chennai, India",
      status: "In Transit",
      updated: "30 Sep 2026",
    },
  ]);

  const [formData, setFormData] = useState({
    category: "Local",
    subType: "Porter",
    carrier: "Porter",
    crnNo: "",
    ewayBillNo: "",
    deliveryChallanNo: "",
    trackingNo: "",
    gst: "",
    amount: "",
    origin: "",
    destination: "",
  });

  const mapApiShipment = (s) => ({
    id: s.id,
    category: s.category,
    subType: s.shipment_type || (s.category === "International" ? "International Shipment" : "Carrier / Transporter"),
    carrier: s.name || s.shipment_type || "-",
    crnNo: s.crn_no || "",
    ewayBillNo: s.eway_bill_no || "",
    deliveryChallanNo: s.delivery_challan_no || "",
    trackingNo: s.tracking_id || s.tracking_lr_no || "",
    gst: s.gst || "",
    amount: s.amount ?? "",
    origin: s.origin || "-",
    destination: s.destination || "-",
    status: s.status || "Pending",
    updated: s.updated_at ? new Date(s.updated_at).toLocaleString() : "Just now",
  });

  const loadShipments = async () => {
    setIsLoadingShipments(true);
    try {
      const response = await fetch(`${API_BASE}/api/shipments`);
      if (!response.ok) throw new Error("Failed to load shipments");
      const data = await response.json();
      setShipments(data.map(mapApiShipment));
      setApiError("");
    } catch (error) {
      setApiError("Central database is not connected. Start the backend to use shared company data.");
    } finally {
      setIsLoadingShipments(false);
    }
  };

  useEffect(() => {
    loadShipments();
  }, []);

  const filteredShipments =
    activeCategory === "All"
      ? shipments
      : shipments.filter(
          (item) => item.category === activeCategory
        );

  const totalShipments = shipments.length;

  const localCount = shipments.filter(
    (item) => item.category === "Local"
  ).length;

  const domesticCount = shipments.filter(
    (item) => item.category === "Domestic"
  ).length;

  const internationalCount = shipments.filter(
    (item) => item.category === "International"
  ).length;

  const inTransitCount = shipments.filter(
    (item) => item.status === "In Transit"
  ).length;

  const deliveredCount = shipments.filter(
    (item) => item.status === "Delivered"
  ).length;

  const pendingCount = shipments.filter(
    (item) => item.status === "Pending"
  ).length;

  const getTrackingReference = (item) => {
    if (item.category === "Local" && item.subType === "Porter") {
      return item.crnNo || "";
    }

    if (item.category === "Local" || item.category === "Domestic") {
      return item.trackingNo || "";
    }

    return item.trackingNo || "";
  };

  const isTrackingStarted = (item) => Boolean(getTrackingReference(item));

  const trackingSteps = [
    "Pending",
    "In Transit",
    "Out for Delivery",
    "Delivered",
    "Completed",
  ];

  const getTrackingStepIndex = (status) => {
    const index = trackingSteps.indexOf(status);
    return index >= 0 ? index : 0;
  };

  // ==========================================
  // RESET FORM
  // ==========================================

  // ==========================================
  // EXCEL IMPORT
  // ==========================================

  const normalizeHeader = (value) =>
    String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/[\s_./-]+/g, "");

  const getCell = (row, aliases) => {
    const keys = Object.keys(row);
    const wanted = aliases.map(normalizeHeader);
    const key = keys.find((item) =>
      wanted.includes(normalizeHeader(item))
    );

    return key ? String(row[key] ?? "").trim() : "";
  };

  const makeImportedShipment = ({
    category,
    subType,
    carrier,
    crnNo,
    ewayBillNo,
    deliveryChallanNo,
    trackingNo,
    gst,
    amount,
    origin,
    destination,
  }) => ({
    id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
    category,
    subType,
    carrier: carrier || "-",
    crnNo: crnNo || "",
    ewayBillNo: ewayBillNo || "",
    deliveryChallanNo: deliveryChallanNo || "",
    trackingNo: trackingNo || "",
    gst: gst || "",
    amount: amount ?? "",
    origin: origin || "-",
    destination: destination || "-",
    status: "Pending",
    updated: "Just now",
  });

  const validateImportedRow = (category, row) => {
    const carrier = getCell(row, [
      "Carrier",
      "Transporter",
      "Carrier / Transporter",
      "Shipping Line",
      "Freight Forwarder",
      "Shipping Line / Freight Forwarder",
    ]);

    const crnNo = getCell(row, [
      "CRN",
      "CRN No",
      "CRN No.",
      "CRN Number",
    ]);

    const ewayBillNo = getCell(row, [
      "E-Way Bill",
      "E-Way Bill No",
      "E-Way Bill No.",
      "Eway Bill",
      "Eway Bill No",
      "EWB",
    ]);

    const deliveryChallanNo = getCell(row, [
      "Delivery Challan",
      "Delivery Challan No",
      "Delivery Challan No.",
      "DC",
      "DC No",
    ]);

    const trackingNo = getCell(row, [
      "Tracking ID",
      "Tracking",
      "Tracking No",
      "Tracking No.",
      "LR No",
      "LR Number",
      "BL No",
      "B/L No",
      "Bill of Lading",
      "Container No",
      "Container Number",
    ]);

    const origin = getCell(row, ["Origin", "Pickup Location"]);
    const destination = getCell(row, [
      "Destination",
      "Delivery Location",
    ]);

    if (category === "Local") {
      const hasCrn = Boolean(crnNo);
      const hasCarrierDocs =
        Boolean(ewayBillNo) && Boolean(deliveryChallanNo);

      if (hasCrn && !ewayBillNo && !deliveryChallanNo) {
        return {
          valid: true,
          shipment: makeImportedShipment({
            category: "Local",
            subType: "Porter",
            carrier: carrier || "Porter",
            crnNo,
            ewayBillNo: "",
            deliveryChallanNo: "",
            trackingNo,
            origin,
            destination,
          }),
        };
      }

      if (!crnNo && hasCarrierDocs) {
        return {
          valid: true,
          shipment: makeImportedShipment({
            category: "Local",
            subType: "Carrier / Transporter",
            carrier,
            crnNo: "",
            ewayBillNo,
            deliveryChallanNo,
            trackingNo,
            origin,
            destination,
          }),
        };
      }

      if (!crnNo && ewayBillNo && !deliveryChallanNo) {
        return {
          valid: false,
          reason:
            "Local Carrier / Transporter requires both E-Way Bill No. and Delivery Challan No.",
        };
      }

      if (!crnNo && !ewayBillNo && deliveryChallanNo) {
        return {
          valid: false,
          reason:
            "Local Carrier / Transporter requires both E-Way Bill No. and Delivery Challan No.",
        };
      }

      if (crnNo && (ewayBillNo || deliveryChallanNo)) {
        return {
          valid: false,
          reason:
            "Local row is mismatched. Porter uses CRN only; Carrier / Transporter uses E-Way Bill No. + Delivery Challan No.",
        };
      }

      return {
        valid: false,
        reason:
          "Local row is missing mandatory data. Enter CRN No. for Porter, or E-Way Bill No. + Delivery Challan No. for Carrier / Transporter.",
      };
    }

    if (category === "Domestic") {
      if (!ewayBillNo) {
        return {
          valid: false,
          reason: "Domestic shipment requires E-Way Bill No.",
        };
      }

      if (!deliveryChallanNo) {
        return {
          valid: false,
          reason: "Domestic shipment requires Delivery Challan No.",
        };
      }

      return {
        valid: true,
        shipment: makeImportedShipment({
          category: "Domestic",
          subType: "Carrier / Transporter",
          carrier,
          crnNo: "",
          ewayBillNo,
          deliveryChallanNo,
          trackingNo,
          gst,
          amount,
          origin,
          destination,
        }),
      };
    }

    if (category === "International") {
      if (!trackingNo) {
        return {
          valid: false,
          reason: "International shipment requires Tracking ID.",
        };
      }

      return {
        valid: true,
        shipment: makeImportedShipment({
          category: "International",
          subType: "International Shipment",
          carrier,
          crnNo: "",
          ewayBillNo: "",
          deliveryChallanNo: "",
          trackingNo,
          gst,
          amount,
          origin,
          destination,
        }),
      };
    }

    return {
      valid: false,
      reason: "Unknown sheet category.",
    };
  };

  const handleDownloadTemplate = () => {
    const link = document.createElement("a");
    link.href = "/Logistics_Shipment_Upload_Template.xlsx";
    link.download = "Logistics_Shipment_Upload_Template.xlsx";
    document.body.appendChild(link);
    link.click();
    link.remove();
  };

  const handleExcelImport = async (e) => {
    const file = e.target.files?.[0];

    if (!file) {
      return;
    }

    setIsImporting(true);
    setImportResult(null);

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });

      const requiredSheets = [
        "Local",
        "Domestic",
        "International",
      ];

      const missingSheets = requiredSheets.filter(
        (sheetName) => !workbook.SheetNames.includes(sheetName)
      );

      if (missingSheets.length > 0) {
        setImportResult({
          success: false,
          message: `Excel rejected. Missing sheet(s): ${missingSheets.join(
            ", "
          )}. The workbook must contain Local, Domestic and International sheets.`,
          rejected: [],
        });
        return;
      }

      const extraSheets = workbook.SheetNames.filter(
        (sheetName) => !requiredSheets.includes(sheetName)
      );

      if (extraSheets.length > 0) {
        setImportResult({
          success: false,
          message: `Excel rejected. Extra sheet(s) found: ${extraSheets.join(
            ", "
          )}. Use only Local, Domestic and International sheets.`,
          rejected: [],
        });
        return;
      }

      const accepted = [];
      const rejected = [];

      requiredSheets.forEach((category) => {
        const sheet = workbook.Sheets[category];

        const rows = XLSX.utils.sheet_to_json(sheet, {
          defval: "",
          raw: false,
        });

        if (rows.length === 0) {
          rejected.push({
            sheet: category,
            row: "-",
            reason: "Sheet has no shipment rows.",
          });
          return;
        }

        rows.forEach((row, index) => {
          const result = validateImportedRow(category, row);

          if (result.valid) {
            accepted.push(result.shipment);
          } else {
            rejected.push({
              sheet: category,
              row: index + 2,
              reason: result.reason,
            });
          }
        });
      });

      if (rejected.length > 0) {
        setImportResult({
          success: false,
          message:
            "Excel rejected. No imported shipment was added because one or more rows failed validation. Correct the rejected rows and import again.",
          rejected,
        });
        return;
      }

      const saved = [];
      for (const shipment of accepted) {
        const response = await fetch(`${API_BASE}/api/shipments`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            category: shipment.category,
            shipment_type: shipment.subType,
            name: shipment.carrier,
            crn_no: shipment.crnNo,
            eway_bill_no: shipment.ewayBillNo,
            delivery_challan_no: shipment.deliveryChallanNo,
            tracking_lr_no: shipment.trackingNo,
            tracking_id: shipment.category === "International" ? shipment.trackingNo : "",
            gst: shipment.gst,
            amount: shipment.amount,
            origin: shipment.origin,
            destination: shipment.destination,
          }),
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.errors?.join(" ") || data.message || "Import save failed.");
        saved.push(mapApiShipment(data));
      }

      setShipments((prev) => [...saved, ...prev]);
      setApiError("");
      setImportResult({
        success: true,
        message: `${saved.length} shipment(s) imported successfully to the central database.`,
        rejected: [],
      });
    } catch (error) {
      setImportResult({
        success: false,
        message:
          "Excel rejected. The file could not be read. Please use a valid .xlsx or .xls file with Local, Domestic and International sheets.",
        rejected: [],
      });
    } finally {
      setIsImporting(false);
      e.target.value = "";
    }
  };

  const resetForm = () => {
    setFormData({
      category: "Local",
      subType: "Porter",
      carrier: "Porter",
      crnNo: "",
      ewayBillNo: "",
      deliveryChallanNo: "",
      trackingNo: "",
      gst: "",
      amount: "",
      origin: "",
      destination: "",
    });
  };

  // ==========================================
  // CATEGORY CHANGE
  // ==========================================

  const handleCategoryChange = (category) => {
    if (category === "Local") {
      setFormData({
        category: "Local",
        subType: "Porter",
        carrier: "Porter",
        crnNo: "",
        ewayBillNo: "",
        deliveryChallanNo: "",
        trackingNo: "",
        origin: "",
        destination: "",
      });

      return;
    }

    if (category === "Domestic") {
      setFormData({
        category: "Domestic",
        subType: "Carrier / Transporter",
        carrier: "",
        crnNo: "",
        ewayBillNo: "",
        deliveryChallanNo: "",
        trackingNo: "",
        origin: "",
        destination: "",
      });

      return;
    }

    if (category === "International") {
      setFormData({
        category: "International",
        subType: "International Shipment",
        carrier: "",
        crnNo: "",
        ewayBillNo: "",
        deliveryChallanNo: "",
        trackingNo: "",
        origin: "",
        destination: "",
      });
    }
  };

  // ==========================================
  // LOCAL TYPE CHANGE
  // ==========================================

  const handleLocalTypeChange = (subType) => {
    if (subType === "Porter") {
      setFormData({
        ...formData,
        subType: "Porter",
        carrier: "Porter",
        crnNo: "",
        ewayBillNo: "",
        deliveryChallanNo: "",
        trackingNo: "",
      });

      return;
    }

    setFormData({
      ...formData,
      subType: "Carrier / Transporter",
      carrier: "",
      crnNo: "",
      ewayBillNo: "",
      deliveryChallanNo: "",
      trackingNo: "",
    });
  };

  // ==========================================
  // VALIDATION
  // ==========================================

  const validateForm = () => {
    if (formData.category === "Local") {
      if (formData.subType === "Porter") {
        if (!formData.crnNo.trim()) {
          alert("CRN No. is mandatory for Porter.");
          return false;
        }
      }

      if (formData.subType === "Carrier / Transporter") {
        if (!formData.ewayBillNo.trim()) {
          alert("E-Way Bill No. is mandatory.");
          return false;
        }

        if (!formData.deliveryChallanNo.trim()) {
          alert("Delivery Challan No. is mandatory.");
          return false;
        }
      }
    }

    if (formData.category === "Domestic") {
      if (!formData.ewayBillNo.trim()) {
        alert("E-Way Bill No. is mandatory for Domestic shipment.");
        return false;
      }

      if (!formData.deliveryChallanNo.trim()) {
        alert(
          "Delivery Challan No. is mandatory for Domestic shipment."
        );
        return false;
      }
    }

    if (formData.category === "International") {
      if (!formData.trackingNo.trim()) {
        alert(
          "Tracking ID is mandatory for International shipment."
        );
        return false;
      }
    }

    return true;
  };

  // ==========================================
  // ADD SHIPMENT
  // ==========================================

  const handleAddShipment = async (e) => {
    e.preventDefault();

    const isValid = validateForm();

    if (!isValid) {
      return;
    }

    const newShipment = {
      id: Date.now(),
      category: formData.category,
      subType: formData.subType,
      carrier: formData.carrier || "-",
      crnNo: formData.crnNo || "",
      ewayBillNo: formData.ewayBillNo || "",
      deliveryChallanNo: formData.deliveryChallanNo || "",
      trackingNo: formData.trackingNo || "",
      gst: formData.gst || "",
      amount: formData.amount || "",
      origin: formData.origin || "-",
      destination: formData.destination || "-",
      status: "Pending",
      updated: "Just now",
    };

    try {
      const response = await fetch(`${API_BASE}/api/shipments`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          category: newShipment.category,
          shipment_type: newShipment.subType,
          name: newShipment.carrier,
          crn_no: newShipment.crnNo,
          eway_bill_no: newShipment.ewayBillNo,
          delivery_challan_no: newShipment.deliveryChallanNo,
          tracking_lr_no: newShipment.trackingNo,
          tracking_id: newShipment.category === "International" ? newShipment.trackingNo : "",
          gst: newShipment.gst,
          amount: newShipment.amount,
          origin: newShipment.origin,
          destination: newShipment.destination,
        }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.errors?.join(" ") || data.message || "Failed to save shipment.");
      setShipments((prev) => [mapApiShipment(data), ...prev]);
      setApiError("");
      resetForm();
      setShowAddForm(false);
    } catch (error) {
      setApiError(error.message || "Central database connection failed.");
      alert(error.message || "Central database connection failed. Start the backend and try again.");
    }
  };

  return (
    <div className="app">

      {/* ==========================================
          HEADER
      ========================================== */}

      <header className="header">

        <div>
          <h1>Logistics Tracking System</h1>

          <p>
            Central shipment tracking dashboard
          </p>
        </div>

        <div className="header-badge">
          ● SYSTEM ONLINE
        </div>

      </header>


      {/* ==========================================
          MAIN
      ========================================== */}

      <main className="container">

        {/* ==========================================
            DASHBOARD HEADING
        ========================================== */}

        <section className="dashboard-heading">

          <div>

            <h2>Shipment Dashboard</h2>

            <p>
              Track and manage Local, Domestic and
              International shipments
            </p>

          </div>

          <button
            className="add-shipment-btn"
            onClick={() => setShowAddForm(true)}
          >
            + Add Shipment
          </button>

        </section>


        {/* ==========================================
            CATEGORY CARDS
        ========================================== */}

        <section className="category-grid">

          <button
            className={`category-card ${
              activeCategory === "Local"
                ? "active"
                : ""
            }`}
            onClick={() =>
              setActiveCategory("Local")
            }
          >

            <div className="category-icon">
              📦
            </div>

            <div>

              <span>LOCAL</span>

              <strong>
                {localCount}
              </strong>

              <small>
                Shipments
              </small>

            </div>

          </button>


          <button
            className={`category-card ${
              activeCategory === "Domestic"
                ? "active"
                : ""
            }`}
            onClick={() =>
              setActiveCategory("Domestic")
            }
          >

            <div className="category-icon">
              🚚
            </div>

            <div>

              <span>DOMESTIC</span>

              <strong>
                {domesticCount}
              </strong>

              <small>
                Shipments
              </small>

            </div>

          </button>


          <button
            className={`category-card ${
              activeCategory === "International"
                ? "active"
                : ""
            }`}
            onClick={() =>
              setActiveCategory("International")
            }
          >

            <div className="category-icon">
              🚢
            </div>

            <div>

              <span>INTERNATIONAL</span>

              <strong>
                {internationalCount}
              </strong>

              <small>
                Shipments
              </small>

            </div>

          </button>

        </section>


        {/* ==========================================
            SUMMARY
        ========================================== */}

        <section className="summary-card">

          <div>

            <span>
              Total Shipments
            </span>

            <strong>
              {totalShipments}
            </strong>

          </div>


          <div>

            <span>
              In Transit
            </span>

            <strong>
              {inTransitCount}
            </strong>

          </div>


          <div>

            <span>
              Delivered
            </span>

            <strong>
              {deliveredCount}
            </strong>

          </div>


          <div>

            <span>
              Pending
            </span>

            <strong>
              {pendingCount}
            </strong>

          </div>

        </section>


        {/* ==========================================
            SHIPMENT TABLE
        ========================================== */}

        <section className="table-card">

          <div className="section-title">

            <div>

              <h2>
                {activeCategory === "All"
                  ? "All Shipments"
                  : `${activeCategory} Shipments`}
              </h2>

              <p>
                Shipment tracking information
              </p>

            </div>


            <button
              className="all-btn"
              onClick={() =>
                setActiveCategory("All")
              }
            >
              View All
            </button>

          </div>


          <div className="table-wrapper">

            <table>

              <thead>

                <tr>

                  <th>
                    Type
                  </th>

                  <th>
                    Mode
                  </th>

                  <th>
                    Carrier
                  </th>

                  <th>
                    CRN / E-Way / DC / Tracking
                  </th>

                  <th>
                    Origin
                  </th>

                  <th>
                    Destination
                  </th>

                  <th>
                    Status
                  </th>

                  <th>
                    Last Update
                  </th>

                </tr>

              </thead>


              <tbody>

                {filteredShipments.map((item) => (

                  <tr key={item.id}>

                    <td>

                      <span
                        className={`type-badge ${item.category.toLowerCase()}`}
                      >
                        {item.category}
                      </span>

                    </td>


                    <td>
                      <strong>
                        {item.subType}
                      </strong>
                    </td>


                    <td>
                      <strong>
                        {item.carrier}
                      </strong>
                    </td>


                    <td>

                      {item.category === "Local" &&
                        item.subType === "Porter" &&
                        item.crnNo && (
                          <div>
                            <small>
                              CRN:
                            </small>{" "}
                            {item.crnNo}
                          </div>
                        )}


                      {item.ewayBillNo && (
                        <div>
                          <small>
                            E-Way:
                          </small>{" "}
                          {item.ewayBillNo}
                        </div>
                      )}


                      {item.deliveryChallanNo && (
                        <div>
                          <small>
                            DC:
                          </small>{" "}
                          {item.deliveryChallanNo}
                        </div>
                      )}


                      {item.trackingNo && (
                        <div>
                          <small>
                            Tracking:
                          </small>{" "}
                          {item.trackingNo}
                        </div>
                      )}

                    </td>


                    <td>
                      {item.origin}
                    </td>


                    <td>
                      {item.destination}
                    </td>


                    <td>

                      <span
                        className={`status ${item.status
                          .toLowerCase()
                          .replace(" ", "-")}`}
                      >
                        {item.status}
                      </span>

                      <button
                        type="button"
                        onClick={() => setSelectedTrackingShipment(item)}
                        style={{
                          marginTop: "8px",
                          border: "1px solid #d8e5ec",
                          background: "#ffffff",
                          color: "#1f5d7a",
                          borderRadius: "7px",
                          padding: "5px 9px",
                          fontSize: "12px",
                          fontWeight: 700,
                          cursor: "pointer",
                        }}
                      >
                        Track
                      </button>

                    </td>


                    <td>
                      {item.updated}
                    </td>

                  </tr>

                ))}

              </tbody>

            </table>

          </div>

        </section>


        {selectedTrackingShipment && (
          <div className="modal-overlay">
            <div className="modal-card" style={{ maxWidth: "620px" }}>
              <div className="modal-header">
                <div>
                  <h2>Shipment Tracking</h2>
                  <p>Tracking reference and current shipment status</p>
                </div>
                <button
                  className="close-btn"
                  onClick={() => setSelectedTrackingShipment(null)}
                >
                  ×
                </button>
              </div>

              <div style={{
                background: "#f4f9fc",
                border: "1px solid #d9e7ef",
                borderRadius: "10px",
                padding: "15px",
                marginBottom: "18px",
              }}>
                <div style={{ fontWeight: 800, marginBottom: "6px" }}>
                  {selectedTrackingShipment.category} • {selectedTrackingShipment.subType}
                </div>
                <div style={{ color: "#526b7a", fontSize: "13px" }}>
                  Reference: {getTrackingReference(selectedTrackingShipment) || "Not available yet"}
                </div>
              </div>

              {!isTrackingStarted(selectedTrackingShipment) ? (
                <div style={{
                  padding: "16px",
                  borderRadius: "10px",
                  background: "#fff8e8",
                  border: "1px solid #f0dfb0",
                  color: "#795b13",
                  lineHeight: 1.6,
                }}>
                  Tracking has not started yet. The required tracking reference is not available.
                  {selectedTrackingShipment.category !== "International" && " Add the Tracking / LR No. to start tracking."}
                </div>
              ) : (
                <div>
                  <div style={{ fontWeight: 800, marginBottom: "14px" }}>
                    Current status: {selectedTrackingShipment.status}
                  </div>

                  <div style={{ display: "grid", gap: "10px" }}>
                    {trackingSteps.map((step, index) => {
                      const currentIndex = getTrackingStepIndex(selectedTrackingShipment.status);
                      const completed = index <= currentIndex;
                      return (
                        <div key={step} style={{
                          display: "flex",
                          alignItems: "center",
                          gap: "10px",
                          padding: "10px 12px",
                          borderRadius: "8px",
                          background: completed ? "#f0f8f4" : "#f7f9fa",
                          border: "1px solid #e1e9ed",
                        }}>
                          <span style={{
                            width: "22px",
                            height: "22px",
                            borderRadius: "50%",
                            display: "inline-flex",
                            alignItems: "center",
                            justifyContent: "center",
                            fontSize: "12px",
                            fontWeight: 800,
                            background: completed ? "#2f8f63" : "#dce5ea",
                            color: completed ? "#fff" : "#6d7d86",
                          }}>
                            {completed ? "✓" : index + 1}
                          </span>
                          <span style={{ fontWeight: completed ? 700 : 500 }}>{step}</span>
                        </div>
                      );
                    })}
                  </div>

                  <p style={{ marginTop: "15px", fontSize: "12px", color: "#6b7d88" }}>
                    Status updates are ready to be connected to the respective carrier API.
                  </p>
                </div>
              )}
            </div>
          </div>
        )}

        {/* ==========================================
            ADD SHIPMENT MODAL
        ========================================== */}

        {showImportForm && (
          <div className="modal-overlay">
            <div className="modal-card">
              <div className="modal-header">
                <div>
                  <h2>Import Data</h2>
                  <p>
                    Upload completed Excel data in bulk or download the official template.
                  </p>
                </div>

                <button
                  className="close-btn"
                  onClick={() => {
                    setImportResult(null);
                    setShowImportForm(false);
                  }}
                >
                  ×
                </button>
              </div>

              <div
                style={{
                  background: "#f4f9fc",
                  border: "1px solid #d9e7ef",
                  borderRadius: "10px",
                  padding: "15px",
                  marginBottom: "18px",
                  fontSize: "13px",
                  lineHeight: "1.7",
                  color: "#405c70",
                }}
              >
                <strong>Excel conditions:</strong>
                <br />
                • One workbook must contain exactly 3 sheets: Local, Domestic and International.
                <br />
                • Local → Porter: CRN No. is mandatory.
                <br />
                • Local → Carrier / Transporter: E-Way Bill No. + Delivery Challan No. are mandatory.
                <br />
                • Domestic: E-Way Bill No. + Delivery Challan No. are mandatory.
                <br />
                • International: Tracking ID is mandatory.
                <br />
                • Optional fields can be blank.
                <br />
                • If any mandatory data is missing or mismatched, the Excel upload is rejected and no shipment is added.
              </div>

              <div
                style={{
                  display: "grid",
                  gap: "12px",
                }}
              >
                <label
                  className="add-shipment-btn"
                  style={{
                    display: "inline-flex",
                    justifyContent: "center",
                    width: "100%",
                    cursor: isImporting ? "not-allowed" : "pointer",
                    opacity: isImporting ? 0.7 : 1,
                    boxSizing: "border-box",
                  }}
                >
                  {isImporting ? "Validating Excel..." : "↑ Upload Data"}
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={handleExcelImport}
                    disabled={isImporting}
                    style={{ display: "none" }}
                  />
                </label>

                <button
                  type="button"
                  className="cancel-btn"
                  style={{
                    width: "100%",
                    minHeight: "46px",
                    fontWeight: 700,
                  }}
                  onClick={handleDownloadTemplate}
                  disabled={isImporting}
                >
                  ↓ Download Excel Template
                </button>
              </div>

              {importResult && (
                <div
                  style={{
                    marginTop: "18px",
                    padding: "14px",
                    borderRadius: "10px",
                    background: importResult.success
                      ? "#def7e8"
                      : "#fff1d5",
                    color: importResult.success
                      ? "#15803d"
                      : "#8a5200",
                    fontSize: "13px",
                    lineHeight: "1.6",
                  }}
                >
                  <strong>
                    {importResult.success ? "Success" : "Rejected"}
                  </strong>
                  <br />
                  {importResult.message}

                  {importResult.rejected.length > 0 && (
                    <div style={{ marginTop: "10px" }}>
                      <strong>Rejected rows:</strong>

                      {importResult.rejected.map((item, index) => (
                        <div
                          key={`${item.sheet}-${item.row}-${index}`}
                        >
                          {item.sheet} — Row {item.row}: {item.reason}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}

              <div className="form-actions">
                <button
                  type="button"
                  className="cancel-btn"
                  onClick={() => {
                    setImportResult(null);
                    setShowImportForm(false);
                  }}
                >
                  Close
                </button>
              </div>
            </div>
          </div>
        )}

        {showAddForm && (

          <div className="modal-overlay">

            <div className="modal-card">

              {/* HEADER */}

              <div className="modal-header">

                <div>

                  <h2>
                    Add Shipment
                  </h2>

                  <p>
                    Select shipment category and
                    enter required details
                  </p>

                  <button
                    type="button"
                    className="add-shipment-btn"
                    style={{ marginTop: "12px" }}
                    onClick={() => {
                      setShowAddForm(false);
                      setImportResult(null);
                      setShowImportForm(true);
                    }}
                  >
                    ⇩ Import Data
                  </button>

                </div>


                <button
                  className="close-btn"
                  onClick={() => {
                    resetForm();
                    setShowAddForm(false);
                  }}
                >
                  ×
                </button>

              </div>


              <form
                onSubmit={handleAddShipment}
              >

                {/* ==========================================
                    CATEGORY TABS
                ========================================== */}

                <div className="shipment-tabs">

                  <button
                    type="button"
                    className={
                      formData.category === "Local"
                        ? "shipment-tab active"
                        : "shipment-tab"
                    }
                    onClick={() =>
                      handleCategoryChange("Local")
                    }
                  >
                    Local
                  </button>


                  <button
                    type="button"
                    className={
                      formData.category === "Domestic"
                        ? "shipment-tab active"
                        : "shipment-tab"
                    }
                    onClick={() =>
                      handleCategoryChange("Domestic")
                    }
                  >
                    Domestic
                  </button>


                  <button
                    type="button"
                    className={
                      formData.category === "International"
                        ? "shipment-tab active"
                        : "shipment-tab"
                    }
                    onClick={() =>
                      handleCategoryChange(
                        "International"
                      )
                    }
                  >
                    International
                  </button>

                </div>


                {/* ==========================================
                    LOCAL
                ========================================== */}

                {formData.category === "Local" && (

                  <>

                    <div className="form-group full-width">

                      <label>
                        Local Shipment Type
                      </label>

                      <select
                        value={formData.subType}
                        onChange={(e) =>
                          handleLocalTypeChange(
                            e.target.value
                          )
                        }
                      >

                        <option value="Porter">
                          Porter
                        </option>

                        <option value="Carrier / Transporter">
                          Carrier / Transporter
                        </option>

                      </select>

                    </div>


                    {/* PORTER */}

                    {formData.subType ===
                      "Porter" && (

                      <div className="form-grid">

                        <div className="form-group">

                          <label>
                            CRN No. *
                          </label>

                          <input
                            type="text"
                            placeholder="Enter Porter CRN No."
                            value={formData.crnNo}
                            onChange={(e) =>
                              setFormData({
                                ...formData,
                                crnNo:
                                  e.target.value,
                              })
                            }
                          />

                        </div>


                        <div className="form-group">

                          <label>
                            Origin
                          </label>

                          <input
                            type="text"
                            placeholder="Optional"
                            value={formData.origin}
                            onChange={(e) =>
                              setFormData({
                                ...formData,
                                origin:
                                  e.target.value,
                              })
                            }
                          />

                        </div>


                        <div className="form-group">

                          <label>
                            Destination
                          </label>

                          <input
                            type="text"
                            placeholder="Optional"
                            value={formData.destination}
                            onChange={(e) =>
                              setFormData({
                                ...formData,
                                destination:
                                  e.target.value,
                              })
                            }
                          />

                        </div>

                      </div>

                    )}


                    {/* LOCAL CARRIER */}

                    {formData.subType ===
                      "Carrier / Transporter" && (

                      <div className="form-grid">

                        <div className="form-group">

                          <label>
                            Carrier / Transporter
                          </label>

                          <input
                            type="text"
                            placeholder="Enter transporter name"
                            value={formData.carrier}
                            onChange={(e) =>
                              setFormData({
                                ...formData,
                                carrier:
                                  e.target.value,
                              })
                            }
                          />

                        </div>


                        <div className="form-group">

                          <label>
                            E-Way Bill No. *
                          </label>

                          <input
                            type="text"
                            placeholder="Enter E-Way Bill No."
                            value={
                              formData.ewayBillNo
                            }
                            onChange={(e) =>
                              setFormData({
                                ...formData,
                                ewayBillNo:
                                  e.target.value,
                              })
                            }
                          />

                        </div>


                        <div className="form-group">

                          <label>
                            Delivery Challan No. *
                          </label>

                          <input
                            type="text"
                            placeholder="Enter Delivery Challan No."
                            value={
                              formData.deliveryChallanNo
                            }
                            onChange={(e) =>
                              setFormData({
                                ...formData,
                                deliveryChallanNo:
                                  e.target.value,
                              })
                            }
                          />

                        </div>


                        <div className="form-group">

                          <label>
                            Tracking / LR No.
                          </label>

                          <input
                            type="text"
                            placeholder="Optional - update later"
                            value={
                              formData.trackingNo
                            }
                            onChange={(e) =>
                              setFormData({
                                ...formData,
                                trackingNo:
                                  e.target.value,
                              })
                            }
                          />

                        </div>


                        <div className="form-group">

                          <label>
                            Origin
                          </label>

                          <input
                            type="text"
                            placeholder="Optional"
                            value={formData.origin}
                            onChange={(e) =>
                              setFormData({
                                ...formData,
                                origin:
                                  e.target.value,
                              })
                            }
                          />

                        </div>


                        <div className="form-group">

                          <label>
                            Destination
                          </label>

                          <input
                            type="text"
                            placeholder="Optional"
                            value={
                              formData.destination
                            }
                            onChange={(e) =>
                              setFormData({
                                ...formData,
                                destination:
                                  e.target.value,
                              })
                            }
                          />

                        </div>

                      </div>

                    )}

                  </>

                )}


                {/* ==========================================
                    DOMESTIC
                ========================================== */}

                {formData.category === "Domestic" && (

                  <div className="form-grid">

                    <div className="form-group">

                      <label>
                        Carrier / Transporter
                      </label>

                      <input
                        type="text"
                        placeholder="Optional"
                        value={formData.carrier}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            carrier:
                              e.target.value,
                          })
                        }
                      />

                    </div>


                    <div className="form-group">

                      <label>
                        E-Way Bill No. *
                      </label>

                      <input
                        type="text"
                        placeholder="Enter E-Way Bill No."
                        value={
                          formData.ewayBillNo
                        }
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            ewayBillNo:
                              e.target.value,
                          })
                        }
                      />

                    </div>


                    <div className="form-group">

                      <label>
                        Delivery Challan No. *
                      </label>

                      <input
                        type="text"
                        placeholder="Enter Delivery Challan No."
                        value={
                          formData.deliveryChallanNo
                        }
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            deliveryChallanNo:
                              e.target.value,
                          })
                        }
                      />

                    </div>


                    <div className="form-group">

                      <label>
                        Tracking / LR No.
                      </label>

                      <input
                        type="text"
                        placeholder="Optional - update later"
                        value={
                          formData.trackingNo
                        }
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            trackingNo:
                              e.target.value,
                          })
                        }
                      />

                    </div>


                    <div className="form-group">

                      <label>
                        Origin
                      </label>

                      <input
                        type="text"
                        placeholder="Optional"
                        value={formData.origin}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            origin:
                              e.target.value,
                          })
                        }
                      />

                    </div>


                    <div className="form-group">

                      <label>
                        Destination
                      </label>

                      <input
                        type="text"
                        placeholder="Optional"
                        value={
                          formData.destination
                        }
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            destination:
                              e.target.value,
                          })
                        }
                      />

                    </div>

                  </div>

                )}


                {/* ==========================================
                    INTERNATIONAL
                ========================================== */}

                {formData.category ===
                  "International" && (

                  <div className="form-grid">

                    <div className="form-group">

                      <label>
                        Shipping Line / Freight Forwarder
                      </label>

                      <input
                        type="text"
                        placeholder="Optional"
                        value={formData.carrier}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            carrier:
                              e.target.value,
                          })
                        }
                      />

                    </div>


                    <div className="form-group">

                      <label>
                        Tracking ID *
                      </label>

                      <input
                        type="text"
                        placeholder="Enter Container No. / B/L No. / Tracking ID"
                        value={
                          formData.trackingNo
                        }
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            trackingNo:
                              e.target.value,
                          })
                        }
                      />

                    </div>


                    <div className="form-group">

                      <label>
                        Origin
                      </label>

                      <input
                        type="text"
                        placeholder="Optional"
                        value={formData.origin}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            origin:
                              e.target.value,
                          })
                        }
                      />

                    </div>


                    <div className="form-group">

                      <label>
                        Destination
                      </label>

                      <input
                        type="text"
                        placeholder="Optional"
                        value={
                          formData.destination
                        }
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            destination:
                              e.target.value,
                          })
                        }
                      />

                    </div>

                  </div>

                )}


                {/* ==========================================
                    ACTION BUTTONS
                ========================================== */}

                <div className="form-actions">

                  <button
                    type="button"
                    className="cancel-btn"
                    onClick={() => {
                      resetForm();
                      setShowAddForm(false);
                    }}
                  >
                    Cancel
                  </button>


                  <button
                    type="submit"
                    className="save-btn"
                  >
                    Add Shipment
                  </button>

                </div>

              </form>

            </div>

          </div>

        )}

      </main>

    </div>
  );
}

export default App;