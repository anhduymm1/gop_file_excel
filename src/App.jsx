import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { useState } from "react";

// Nhường quyền điều khiển cho trình duyệt 1 nhịp, tránh bị "Trang không phản hồi"
const yieldToBrowser = () => new Promise((resolve) => setTimeout(resolve, 0));

// Sau bao nhiêu dòng thì nhường 1 lần. Số nhỏ -> UI mượt hơn nhưng xử lý chậm hơn 1 chút.
const YIELD_EVERY_N_ROWS = 500;

// Lấy giá trị "thuần" (không phải object công thức) từ 1 cell.
// Lý do cần hàm này: khi đọc 1 ô công thức, ExcelJS trả về NGUYÊN OBJECT công thức,
// ví dụ { formula: "...", result: 123 } hoặc, với "shared formula",
// { sharedFormula: "F10", result: 123 } (tham chiếu tới ô công thức "gốc").
// Nếu ta copy nguyên object đó sang sheet TongHop (có số dòng khác hẳn do gộp
// nhiều sheet / bỏ header), tham chiếu sharedFormula sẽ trỏ sai vị trí, và
// ExcelJS sẽ báo lỗi "Shared Formula master must exist above and/or left of
// clone" ngay khi XUẤT file (không phải lúc đọc). Vì vậy ta luôn tách lấy
// GIÁ TRỊ ĐÃ TÍNH (result) ra, không giữ lại công thức, để tránh lỗi này hoàn toàn.
function extractPlainValue(cell, sheetName, warnings) {
  let val;

  try {
    val = cell.value;
  } catch (err) {
    // Ô bị lỗi ngay từ lúc đọc (shared formula hỏng thật sự trong file gốc)
    const model = cell.model;
    let fallback = null;
    let source = "trống (không lấy được giá trị)";

    if (model && typeof model === "object" && "result" in model) {
      fallback = model.result ?? null;
      source = "giá trị đã tính sẵn (cached result)";
    } else {
      try {
        fallback = cell.text ?? null;
        source = "giá trị hiển thị (text)";
      } catch {
        // giữ nguyên fallback = null
      }
    }

    warnings.push({
      sheet: sheetName,
      cell: cell.address,
      error: err?.message || String(err),
      source,
      fallback,
    });

    return fallback;
  }

  // Ô dạng công thức (kể cả shared formula) -> chỉ lấy giá trị đã tính, bỏ công thức
  if (val && typeof val === "object" && !(val instanceof Date)) {
    if ("formula" in val || "sharedFormula" in val) {
      return val.result ?? null;
    }
    // Rich text -> nối thành chuỗi thường
    if (Array.isArray(val.richText)) {
      return val.richText.map((rt) => rt.text).join("");
    }
    // Hyperlink -> lấy phần text hiển thị
    if ("text" in val && "hyperlink" in val) {
      return val.text ?? null;
    }
  }

  return val;
}

// Lấy toàn bộ giá trị 1 dòng, luôn trả về giá trị thuần (đã bỏ công thức),
// có dự phòng khi row.values bị lỗi vì formula hỏng.
function getRowValuesSafe(row, sheetName, warnings) {
  const values = [];
  let i = 0;
  row.eachCell({ includeEmpty: true }, (cell) => {
    values[i] = extractPlainValue(cell, sheetName, warnings);
    i++;
  });
  return values;
}

function App() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState("");
  const [warnings, setWarnings] = useState([]);
  const [startRow, setStartRow] = useState(6);

  const mergeSheets = async () => {
    if (!file) {
      alert("Vui lòng chọn file.");
      return;
    }

    const start = Number(startRow);
    if (!Number.isInteger(start) || start < 1) {
      alert("Dòng bắt đầu không hợp lệ. Vui lòng nhập một số nguyên >= 1.");
      return;
    }

    setLoading(true);
    setProgress("Đang đọc file...");
    setWarnings([]);

    const localWarnings = [];
    let currentContext = "";

    try {
      // Nhường 1 nhịp trước khi bắt đầu để UI kịp render trạng thái loading
      await yieldToBrowser();

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());

      const coloredSheets = [];
      workbook.eachSheet((worksheet) => {
        if (worksheet.properties.tabColor) {
          coloredSheets.push(worksheet);
        }
      });

      if (coloredSheets.length === 0) {
        alert("Không có sheet nào được tô màu.");
        return;
      }

      const outWorkbook = new ExcelJS.Workbook();
      const outSheet = outWorkbook.addWorksheet("TongHop");

      let first = true;
      let totalRowsWritten = 0;

      for (let s = 0; s < coloredSheets.length; s++) {
        const sheet = coloredSheets[s];
        const totalRows = sheet.rowCount;
        let rowsProcessed = 0;

        setProgress(
          `Đang xử lý sheet ${s + 1}/${coloredSheets.length}: "${sheet.name}" (${totalRows} dòng)...`
        );
        await yieldToBrowser();

        for (let rowNumber = start; rowNumber <= totalRows; rowNumber++) {
          const row = sheet.getRow(rowNumber);
          currentContext = `sheet "${sheet.name}", dòng ${rowNumber}`;

          if (!first && rowNumber === start) {
            // chỉ sheet đầu tiên giữ lại dòng header (dòng bắt đầu)
          } else {
            const values = getRowValuesSafe(row, sheet.name, localWarnings);
            outSheet.addRow(values);
            totalRowsWritten++;
          }

          rowsProcessed++;

          if (rowsProcessed % YIELD_EVERY_N_ROWS === 0) {
            setProgress(
              `Sheet ${s + 1}/${coloredSheets.length} "${sheet.name}": ${rowsProcessed}/${totalRows} dòng...`
            );
            await yieldToBrowser();
          }
        }

        first = false;
      }

      setProgress(`Đang tạo file kết quả (${totalRowsWritten} dòng)...`);
      currentContext = "đang tạo file kết quả (writeBuffer)";
      await yieldToBrowser();

      const buffer = await outWorkbook.xlsx.writeBuffer();

      setProgress("Đang tải file về...");
      await yieldToBrowser();

      saveAs(
        new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
        "TongHop.xlsx"
      );

      setWarnings(localWarnings);

      if (localWarnings.length > 0) {
        setProgress(
          `Hoàn tất, nhưng có ${localWarnings.length} ô bị lỗi công thức (xem chi tiết bên dưới).`
        );
      } else {
        setProgress("Hoàn tất!");
      }
    } catch (err) {
      console.error(err);

      // Phân biệt lỗi hết bộ nhớ (thường gặp với file rất nặng) với lỗi khác
      const isMemoryError =
        err instanceof RangeError ||
        /memory|allocation/i.test(err?.message || "");

      const isFormulaError = /shared formula/i.test(err?.message || "");

      const contextSuffix = currentContext
        ? `\n\n(Vị trí xảy ra: ${currentContext})`
        : "";

      if (isFormulaError) {
        alert(
          "File Excel chứa công thức (formula) bị lỗi cấu trúc nội bộ nên không đọc được. " +
          "Thử mở file gốc bằng Excel, chọn toàn bộ vùng dữ liệu, " +
          "Copy -> Paste Special -> Values (dán chỉ giá trị, bỏ công thức), lưu lại rồi gộp lại từ file đó." +
          contextSuffix
        );
      } else if (isMemoryError) {
        alert(
          "File quá nặng khiến trình duyệt hết bộ nhớ khi xử lý. " +
          "Hãy thử: đóng bớt tab khác, dùng trình duyệt trên máy tính (không phải điện thoại), " +
          "hoặc tách file gốc thành nhiều file nhỏ hơn trước khi gộp." +
          contextSuffix
        );
      } else {
        alert(
          "Có lỗi xảy ra: " +
          (err?.message || "không rõ nguyên nhân.") +
          contextSuffix
        );
      }
    } finally {
      setLoading(false);
      setProgress("");
    }
  };

  return (
    <div
      style={{
        padding: 40,
        display: "flex",
        flexDirection: "column",
        gap: 20,
        width: 600,
      }}
    >
      <h2>Merge Colored Sheets</h2>

      <input
        type="file"
        accept=".xlsx"
        onChange={(e) => setFile(e.target.files[0])}
      />

      <label style={{ display: "flex", flexDirection: "column", gap: 4 }}>
        <span>Dữ liệu (kể cả header) bắt đầu từ dòng số:</span>
        <input
          type="number"
          min={1}
          value={startRow}
          onChange={(e) => setStartRow(e.target.value)}
          style={{ padding: "6px 8px", width: 100 }}
        />
      </label>

      <button
        onClick={mergeSheets}
        disabled={!file || loading}
        style={{
          padding: "10px 20px",
          cursor: loading ? "not-allowed" : "pointer",
        }}
      >
        {loading ? "Đang xử lý..." : "Gộp & Tải về"}
      </button>

      {loading && <div>⏳ {progress || "Đang xử lý..."}</div>}

      {!loading && warnings.length > 0 && (
        <div
          style={{
            border: "1px solid #e0a800",
            background: "#fff8e1",
            padding: 12,
            borderRadius: 6,
            fontSize: 13,
          }}
        >
          <strong>
            ⚠️ {warnings.length} ô có công thức lỗi (đã lấy tạm giá trị khác):
          </strong>
          <table
            style={{
              width: "100%",
              marginTop: 8,
              borderCollapse: "collapse",
              fontSize: 12,
            }}
          >
            <thead>
              <tr style={{ textAlign: "left", borderBottom: "1px solid #ccc" }}>
                <th style={{ padding: "4px 6px" }}>Sheet</th>
                <th style={{ padding: "4px 6px" }}>Ô</th>
                <th style={{ padding: "4px 6px" }}>Đã lấy tạm</th>
                <th style={{ padding: "4px 6px" }}>Giá trị</th>
              </tr>
            </thead>
            <tbody>
              {warnings.map((w, idx) => (
                <tr key={idx} style={{ borderBottom: "1px solid #eee" }}>
                  <td style={{ padding: "4px 6px" }}>{w.sheet}</td>
                  <td style={{ padding: "4px 6px", fontWeight: "bold" }}>
                    {w.cell}
                  </td>
                  <td style={{ padding: "4px 6px" }}>{w.source}</td>
                  <td style={{ padding: "4px 6px" }}>
                    {String(w.fallback ?? "(trống)")}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

export default App;