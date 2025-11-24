
import React, { useState, useMemo } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";

// Safely handle default export from CDN if needed
const xlsxModule = (XLSX as any).default || XLSX || {};
const utils = xlsxModule.utils;
const writeFile = xlsxModule.writeFile;

interface InventoryItem {
  id: number;
  title: string;
  price: string;
  condition: string;
  description: string;
  location: string;
  [key: string]: string | number;
}

const inventoryData: InventoryItem[] = [
  { id: 1, title: "Wood Desk", price: "$40", condition: "Good", description: "48\" oak desk", location: "Tucson" },
  { id: 2, title: "Bicycle", price: "$100", condition: "Excellent", description: "New tires", location: "Tucson" },
  { id: 3, title: "Lamp", price: "$15", condition: "Fair", description: "Bronze finish", location: "Tucson" },
  { id: 4, title: "Coffee Table", price: "$25", condition: "Good", description: "Glass top", location: "Phoenix" },
  { id: 5, title: "Bookshelf", price: "$60", condition: "Like New", description: "5-tier wooden", location: "Tucson" },
];

type SortDirection = 'ascending' | 'descending';
type TabType = 'Table' | 'CSV' | 'JSON' | 'SQL' | 'XLSX';

interface SortConfig {
  key: keyof InventoryItem;
  direction: SortDirection;
}

const SortableTable = ({ data }: { data: InventoryItem[] }) => {
  const [filterText, setFilterText] = useState("");
  const [sortConfig, setSortConfig] = useState<SortConfig | null>(null);
  const [activeTab, setActiveTab] = useState<TabType>('Table');

  const sortedAndFilteredData = useMemo(() => {
    let sortableItems = [...data];

    // Filter
    if (filterText) {
      const lowerFilter = filterText.toLowerCase();
      sortableItems = sortableItems.filter((item) =>
        Object.values(item).some((val) =>
          String(val).toLowerCase().includes(lowerFilter)
        )
      );
    }

    // Sort
    if (sortConfig !== null) {
      sortableItems.sort((a, b) => {
        let aValue = a[sortConfig.key];
        let bValue = b[sortConfig.key];

        // Handle Price Sorting (remove $)
        if (sortConfig.key === 'price') {
            const aPrice = parseFloat(String(aValue).replace('$', ''));
            const bPrice = parseFloat(String(bValue).replace('$', ''));
            if (aPrice < bPrice) return sortConfig.direction === 'ascending' ? -1 : 1;
            if (aPrice > bPrice) return sortConfig.direction === 'ascending' ? 1 : -1;
            return 0;
        }

        if (aValue < bValue) {
          return sortConfig.direction === 'ascending' ? -1 : 1;
        }
        if (aValue > bValue) {
          return sortConfig.direction === 'ascending' ? 1 : -1;
        }
        return 0;
      });
    }

    return sortableItems;
  }, [data, filterText, sortConfig]);

  const requestSort = (key: keyof InventoryItem) => {
    let direction: SortDirection = 'ascending';
    if (
      sortConfig &&
      sortConfig.key === key &&
      sortConfig.direction === 'ascending'
    ) {
      direction = 'descending';
    }
    setSortConfig({ key, direction });
  };

  const getSortIndicator = (name: string) => {
    if (!sortConfig || sortConfig.key !== name) return " ↕";
    return sortConfig.direction === 'ascending' ? " ▲" : " ▼";
  };

  // Data format generators
  const generateCSV = (items: InventoryItem[]) => {
    if (items.length === 0) return "";
    const headers = Object.keys(items[0]).filter(k => k !== 'id'); // remove internal ID for cleaner CSV
    const rows = items.map(item => 
      headers.map(key => `"${String(item[key]).replace(/"/g, '""')}"`).join(",")
    );
    return [headers.join(","), ...rows].join("\n");
  };

  const generateSQL = (items: InventoryItem[]) => {
    if (items.length === 0) return "-- No data";
    const tableName = "facebook_listings";
    // Use keys from first item for consistency, excluding internal ID
    const keys = Object.keys(items[0]).filter(k => k !== 'id');
    
    return items.map(item => {
      const columns = keys.join(", ");
      const values = keys.map(key => `'${String(item[key]).replace(/'/g, "''")}'`).join(", ");
      return `INSERT INTO ${tableName} (${columns}) VALUES (${values});`;
    }).join("\n");
  };

  const handleDownload = () => {
    if (activeTab === 'XLSX') {
        if (utils && writeFile) {
            // Prepare data without ID for clean export
            const cleanData = sortedAndFilteredData.map(({id, ...rest}) => rest);
            const ws = utils.json_to_sheet(cleanData);
            const wb = utils.book_new();
            utils.book_append_sheet(wb, ws, "Inventory");
            writeFile(wb, "inventory_export.xlsx");
        } else {
            console.error("XLSX utilities not loaded");
            alert("XLSX library not ready yet. Please try again.");
        }
        return;
    }

    let content = "";
    let type = "";
    let filename = "";

    if (activeTab === 'Table' || activeTab === 'CSV') {
      content = generateCSV(sortedAndFilteredData);
      type = 'text/csv;charset=utf-8;';
      filename = 'inventory_export.csv';
    } else if (activeTab === 'JSON') {
      // Remove ID for clean export
      const cleanData = sortedAndFilteredData.map(({id, ...rest}) => rest);
      content = JSON.stringify(cleanData, null, 2);
      type = 'application/json;charset=utf-8;';
      filename = 'inventory_export.json';
    } else if (activeTab === 'SQL') {
      content = generateSQL(sortedAndFilteredData);
      type = 'text/plain;charset=utf-8;';
      filename = 'inventory_export.sql';
    }

    const blob = new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', filename);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const getDownloadLabel = () => {
    if (activeTab === 'Table') return 'Download CSV';
    return `Download ${activeTab}`;
  }

  return (
    <div className="table-container">
      <div className="action-bar">
        <div className="tabs">
          {(['Table', 'CSV', 'JSON', 'SQL', 'XLSX'] as TabType[]).map((tab) => (
            <button
              key={tab}
              className={`tab-button ${activeTab === tab ? 'active' : ''}`}
              onClick={() => setActiveTab(tab)}
            >
              {tab}
            </button>
          ))}
        </div>
        <button onClick={handleDownload} className="btn-primary" title={`Download data as ${activeTab}`}>
          <span>{getDownloadLabel()}</span>
        </button>
      </div>

      <input
        type="text"
        placeholder="Filter items..."
        value={filterText}
        onChange={(e) => setFilterText(e.target.value)}
        className="search-input"
        aria-label="Filter table rows"
      />

      {activeTab === 'Table' && (
        <div className="table-responsive">
          <table>
            <thead>
              <tr>
                <th onClick={() => requestSort('title')}>Title {getSortIndicator('title')}</th>
                <th onClick={() => requestSort('price')}>Price {getSortIndicator('price')}</th>
                <th onClick={() => requestSort('condition')}>Condition {getSortIndicator('condition')}</th>
                <th onClick={() => requestSort('description')}>Description {getSortIndicator('description')}</th>
                <th onClick={() => requestSort('location')}>Location {getSortIndicator('location')}</th>
              </tr>
            </thead>
            <tbody>
              {sortedAndFilteredData.length > 0 ? (
                sortedAndFilteredData.map((item) => (
                  <tr key={item.id}>
                    <td>{item.title}</td>
                    <td>{item.price}</td>
                    <td>{item.condition}</td>
                    <td>{item.description}</td>
                    <td>{item.location}</td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={5} style={{ textAlign: "center" }}>No items found</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {(activeTab === 'CSV' || activeTab === 'JSON' || activeTab === 'SQL') && (
        <textarea 
          readOnly 
          className="code-view"
          value={
            activeTab === 'CSV' ? generateCSV(sortedAndFilteredData) :
            activeTab === 'JSON' ? JSON.stringify(sortedAndFilteredData.map(({id, ...r}) => r), null, 2) :
            generateSQL(sortedAndFilteredData)
          }
        />
      )}
      
      {activeTab === 'XLSX' && (
        <div className="code-view code-view-placeholder">
          <p>Spreadsheet ready for download. Click the button above.</p>
        </div>
      )}
    </div>
  );
};

const App = () => {
  return (
    <div>
      <h1>One-Page Handout: Why Posting Facebook Marketplace Ads by Spreadsheet Is Valuable</h1>

      <h2>1. What This Method Really Is</h2>
      <p>Think of a spreadsheet as a big organized sheet of paper with tidy boxes.</p>
      
      <SortableTable data={inventoryData} />

      <h2>2. How It Actually Works</h2>
      <ol>
        <li>Fill out the spreadsheet.</li>
        <li>Upload it once to Facebook Marketplace.</li>
        <li>Facebook reads each row and automatically creates a listing for every item.</li>
        <li>All ads appear automatically.</li>
      </ol>

      <h2>3. Why It Saves So Much Time</h2>
      <div className="comparison-box">
        <p><b>Manual posting (one-by-one):</b></p>
        <p>Write Ad → Upload Photos → Enter Price → Enter Description → Post (Repeat 50 times...)</p>
        <hr/>
        <p><b>Spreadsheet upload:</b></p>
        <p>Fill rows → Upload once → Done</p>
      </div>

      <h2>4. Why It Prevents Mistakes</h2>
      <p>Check all details in one place:</p>

      <div className="table-container">
        <table>
          <thead>
            <tr>
              <th>Title</th>
              <th>Price</th>
              <th>Description</th>
              <th>Location</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td className="success-cell">✔ Looks right</td>
              <td className="success-cell">✔ Looks right</td>
              <td className="success-cell">✔ Looks right</td>
              <td className="success-cell">✔ Looks right</td>
            </tr>
          </tbody>
        </table>
      </div>

      <h2>5. Why It Helps You Stay Organized</h2>
      <ul>
        <li>All your items → One sheet</li>
        <li>All your changes → One sheet</li>
        <li>All your ads → Updated at once</li>
      </ul>

      <h2>6. Why Even a Pen-and-Paper Person Can Use This</h2>
      <p>If you’re comfortable writing lists, ledgers, or inventory sheets on paper, you’re already doing the “spreadsheet mindset.”</p>
      <p>The only difference is:</p>
      <ul>
        <li>On paper → you rewrite every ad</li>
        <li>With a spreadsheet → a machine rewrites them for you automatically</li>
      </ul>

      <h2>Simple Analogy</h2>
      <div className="quote-box">
        <p>“Imagine filling out one master catalog instead of handwriting dozens of flyers. Then imagine a helper who takes each line of your catalog and pins it to the bulletin board instantly. That’s what a spreadsheet upload does for Facebook Marketplace.”</p>
      </div>
    </div>
  );
};

const container = document.getElementById("root");
if (container) {
  const root = createRoot(container);
  root.render(<App />);
}
