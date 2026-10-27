/**
 * HỆ THỐNG TỔNG HỢP & GIÁM SÁT TIẾN ĐỘ ÁN HÌNH SỰ - FRONTEND CONTROLLER
 * Viện kiểm sát nhân dân Thành phố Hồ Chí Minh
 */

document.addEventListener("DOMContentLoaded", () => {
    // Elements - Bước 1
    const txtFolderPath = document.getElementById("txtFolderPath");
    const btnBrowseFolder = document.getElementById("btnBrowseFolder");
    const btnBrowseUnitFiles = document.getElementById("btnBrowseUnitFiles");
    const btnClearFolderPath = document.getElementById("btnClearFolderPath");
    const boxSelectedUnitFiles = document.getElementById("boxSelectedUnitFiles");
    const lblSelectedFilesCount = document.getElementById("lblSelectedFilesCount");
    const chipSelectedList = document.getElementById("chipSelectedList");
    const badgeInputMode = document.getElementById("badgeInputMode");

    const txtExistingMaster = document.getElementById("txtExistingMaster");
    const btnBrowseExistingMaster = document.getElementById("btnBrowseExistingMaster");
    const btnClearExistingMaster = document.getElementById("btnClearExistingMaster");
    const spnDays = document.getElementById("spnDays");
    const btnRunStep1 = document.getElementById("btnRunStep1");
    const loadingStep1 = document.getElementById("loadingStep1");
    const resultStep1 = document.getElementById("resultStep1");
    const badgeStep1 = document.getElementById("badgeStep1");

    const boxAppendStats = document.getElementById("boxAppendStats");
    const resExistingCases = document.getElementById("resExistingCases");
    const resNewAdded = document.getElementById("resNewAdded");
    const resSkippedDupes = document.getElementById("resSkippedDupes");
    const lblTotalCases = document.getElementById("lblTotalCases");

    const resTotalFiles = document.getElementById("resTotalFiles");
    const resTotalCases = document.getElementById("resTotalCases");
    const resFilteredCases = document.getElementById("resFilteredCases");
    const pathFileTong = document.getElementById("pathFileTong");
    const pathFileKt = document.getElementById("pathFileKt");
    const timeStep1 = document.getElementById("timeStep1");

    // Elements - Bước 2
    const txtFileTong = document.getElementById("txtFileTong");
    const btnBrowseFileTong = document.getElementById("btnBrowseFileTong");
    const txtFileKt = document.getElementById("txtFileKt");
    const btnBrowseFileKt = document.getElementById("btnBrowseFileKt");
    const btnRunStep2 = document.getElementById("btnRunStep2");
    const loadingStep2 = document.getElementById("loadingStep2");
    const resultStep2 = document.getElementById("resultStep2");

    const kpiTotal = document.getElementById("kpiTotal");
    const kpiKt = document.getElementById("kpiKt");
    const kpiErrors = document.getElementById("kpiErrors");
    const kpiRate = document.getElementById("kpiRate");
    const unitRankingSection = document.getElementById("unitRankingSection");
    const unitTagsContainer = document.getElementById("unitTagsContainer");
    const pathFileMaster = document.getElementById("pathFileMaster");
    const pathFileReport = document.getElementById("pathFileReport");
    const timeStep2 = document.getElementById("timeStep2");

    // Status & Navigation
    const backendStatusText = document.getElementById("backendStatusText");
    const stepIndicator1 = document.getElementById("stepIndicator1");
    const stepIndicator2 = document.getElementById("stepIndicator2");
    const btnQuickOpenFolder = document.getElementById("btnQuickOpenFolder");
    const btnFooterOpenFolder = document.getElementById("btnFooterOpenFolder");
    const toastContainer = document.getElementById("toastContainer");

    let currentWorkingDir = "";

    // 1. Kiểm tra trạng thái Backend
    async function checkBackendConnection() {
        try {
            const res = await fetch("/api/ping");
            if (res.ok) {
                const data = await res.json();
                backendStatusText.textContent = "Backend nhúng: Sẵn sàng (Offline)";
            } else {
                backendStatusText.textContent = "Backend: Phản hồi bất thường";
            }
        } catch (err) {
            backendStatusText.textContent = "Backend: Mất kết nối!";
            showToast("error", "Lỗi kết nối", "Không thể kết nối với Backend máy chủ nội bộ.");
        }
    }
    checkBackendConnection();

    // 2. Hệ thống Toast thông báo
    function showToast(type, title, message, duration = 5000) {
        const toast = document.createElement("div");
        toast.className = `toast ${type}`;

        const icons = {
            success: "✅",
            error: "❌",
            warning: "⚠️"
        };

        toast.innerHTML = `
            <span class="toast-icon">${icons[type] || "ℹ️"}</span>
            <div class="toast-body">
                <div class="toast-title">${title}</div>
                <div class="toast-msg">${message}</div>
            </div>
        `;

        toastContainer.appendChild(toast);
        setTimeout(() => {
            toast.style.opacity = "0";
            toast.style.transform = "translateX(100%)";
            toast.style.transition = "all 0.3s ease";
            setTimeout(() => toast.remove(), 300);
        }, duration);
    }

    // Quản lý danh sách file đơn vị được chọn
    let selectedUnitFiles = [];

    function renderSelectedUnitFiles() {
        if (selectedUnitFiles && selectedUnitFiles.length > 0) {
            boxSelectedUnitFiles.style.display = "flex";
            btnClearFolderPath.style.display = "inline-block";
            lblSelectedFilesCount.textContent = `Đã chọn ${selectedUnitFiles.length} file đơn vị`;
            badgeInputMode.textContent = `📄 Đang chọn ${selectedUnitFiles.length} file đơn vị`;
            badgeInputMode.style.background = "#DCFCE7";
            badgeInputMode.style.color = "#166534";

            chipSelectedList.innerHTML = "";
            selectedUnitFiles.forEach((fpath, idx) => {
                const chip = document.createElement("div");
                chip.className = "chip-item";
                const fname = fpath.split(/[\\/]/).pop();
                chip.innerHTML = `
                    <span>📄 ${fname}</span>
                    <span class="chip-remove" title="Bỏ file này">✕</span>
                `;
                chip.querySelector(".chip-remove").addEventListener("click", () => {
                    selectedUnitFiles.splice(idx, 1);
                    if (selectedUnitFiles.length === 0) {
                        txtFolderPath.value = "";
                    } else {
                        txtFolderPath.value = selectedUnitFiles.join("; ");
                    }
                    renderSelectedUnitFiles();
                });
                chipSelectedList.appendChild(chip);
            });
        } else {
            boxSelectedUnitFiles.style.display = "none";
            if (!txtFolderPath.value.trim()) {
                btnClearFolderPath.style.display = "none";
                badgeInputMode.textContent = "📁 Quét cả thư mục HOẶC 📄 Chọn 1 vài đơn vị";
                badgeInputMode.style.background = "#DBEAFE";
                badgeInputMode.style.color = "#1E40AF";
            }
        }
    }

    // 3. Mở hộp thoại chọn thư mục Windows nguyên bản
    btnBrowseFolder.addEventListener("click", async () => {
        try {
            btnBrowseFolder.disabled = true;
            const res = await fetch("/api/browse-folder", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ current_path: txtFolderPath.value.trim() })
            });
            const data = await res.json();
            if (data.success && data.path) {
                selectedUnitFiles = [];
                txtFolderPath.value = data.path;
                currentWorkingDir = data.path;
                renderSelectedUnitFiles();
                btnClearFolderPath.style.display = "inline-block";
                badgeInputMode.textContent = "📁 Đã chọn thư mục đơn vị";
                badgeInputMode.style.background = "#FEF3C7";
                badgeInputMode.style.color = "#92400E";
            }
        } catch (err) {
            showToast("error", "Lỗi", "Không thể mở hộp thoại chọn thư mục.");
        } finally {
            btnBrowseFolder.disabled = false;
        }
    });

    // 3b. Mở hộp thoại chọn 1 hoặc nhiều file đơn vị (giữ Ctrl để chọn nhiều file)
    btnBrowseUnitFiles.addEventListener("click", async () => {
        try {
            btnBrowseUnitFiles.disabled = true;
            const res = await fetch("/api/browse-files", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    title: "Chọn 1 hoặc nhiều file Excel của các đơn vị (giữ Ctrl để chọn nhiều file)",
                    current_path: currentWorkingDir || ""
                })
            });
            const data = await res.json();
            if (data.success && data.paths && data.paths.length > 0) {
                selectedUnitFiles = data.paths;
                txtFolderPath.value = data.paths.join("; ");
                currentWorkingDir = data.paths[0];
                renderSelectedUnitFiles();
            }
        } catch (err) {
            showToast("error", "Lỗi", "Không thể mở hộp thoại chọn file đơn vị.");
        } finally {
            btnBrowseUnitFiles.disabled = false;
        }
    });

    btnClearFolderPath.addEventListener("click", () => {
        selectedUnitFiles = [];
        txtFolderPath.value = "";
        renderSelectedUnitFiles();
    });

    txtFolderPath.addEventListener("input", () => {
        if (!txtFolderPath.value.trim()) {
            selectedUnitFiles = [];
            renderSelectedUnitFiles();
        } else {
            btnClearFolderPath.style.display = "inline-block";
        }
    });

    function getDirectoryFromPath(p) {
        if (!p || typeof p !== "string") return "";
        const clean = p.split(";")[0].trim().replace(/^["']|["']$/g, "");
        if (!clean) return "";
        if (/\.[a-zA-Z0-9]+$/i.test(clean)) {
            const lastSlash = Math.max(clean.lastIndexOf("/"), clean.lastIndexOf("\\"));
            return lastSlash !== -1 ? clean.substring(0, lastSlash) : "";
        }
        return clean;
    }

    // 4. Mở hộp thoại chọn File Tổng cũ để nạp nối tiếp (Tùy chọn)
    btnBrowseExistingMaster.addEventListener("click", async () => {
        try {
            btnBrowseExistingMaster.disabled = true;
            const targetDir = getDirectoryFromPath(currentWorkingDir || txtFolderPath.value);
            const res = await fetch("/api/browse-file", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    title: "Chọn File Tổng cũ đã có (1_FILE_TONG.xlsx) để nạp thêm vụ mới",
                    current_path: targetDir
                })
            });
            const data = await res.json();
            if (data.success && data.path) {
                txtExistingMaster.value = data.path;
                currentWorkingDir = getDirectoryFromPath(data.path);
                btnClearExistingMaster.style.display = "inline-block";
            }
        } catch (err) {
            showToast("error", "Lỗi", "Không thể mở hộp thoại chọn file tổng cũ.");
        } finally {
            btnBrowseExistingMaster.disabled = false;
        }
    });

    btnClearExistingMaster.addEventListener("click", () => {
        txtExistingMaster.value = "";
        btnClearExistingMaster.style.display = "none";
    });

    txtExistingMaster.addEventListener("input", () => {
        btnClearExistingMaster.style.display = txtExistingMaster.value.trim() ? "inline-block" : "none";
    });

    // 5. Mở hộp thoại chọn file Excel Windows nguyên bản (Bước 2)
    btnBrowseFileTong.addEventListener("click", async () => {
        try {
            btnBrowseFileTong.disabled = true;
            const targetDir = getDirectoryFromPath(currentWorkingDir || txtFolderPath.value);
            const res = await fetch("/api/browse-file", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    title: "Chọn file 1_FILE_TONG.xlsx",
                    current_path: targetDir
                })
            });
            const data = await res.json();
            if (data.success && data.path) {
                txtFileTong.value = data.path;
                currentWorkingDir = getDirectoryFromPath(data.path);
            }
        } catch (err) {
            showToast("error", "Lỗi", "Không thể mở hộp thoại chọn file.");
        } finally {
            btnBrowseFileTong.disabled = false;
        }
    });

    btnBrowseFileKt.addEventListener("click", async () => {
        try {
            btnBrowseFileKt.disabled = true;
            const targetDir = getDirectoryFromPath(currentWorkingDir || txtFolderPath.value);
            const res = await fetch("/api/browse-file", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    title: "Chọn file 2_DANH_SACH_CAN_KIEM_TRA.xlsx đã ghi chú",
                    current_path: targetDir
                })
            });
            const data = await res.json();
            if (data.success && data.path) {
                txtFileKt.value = data.path;
                currentWorkingDir = getDirectoryFromPath(data.path);
            }
        } catch (err) {
            showToast("error", "Lỗi", "Không thể mở hộp thoại chọn file.");
        } finally {
            btnBrowseFileKt.disabled = false;
        }
    });

    // 6. NÚT 1: Chạy tổng hợp và lọc vụ án >= 16 ngày (Hỗ trợ nạp nối tiếp thông minh)
    btnRunStep1.addEventListener("click", async () => {
        const folderPath = txtFolderPath.value.trim();
        if (!folderPath && (!selectedUnitFiles || selectedUnitFiles.length === 0)) {
            showToast("warning", "Chưa chọn nguồn dữ liệu", "Vui lòng bấm 'Chọn thư mục' hoặc 'Chọn file đơn vị' để chọn file báo cáo của các đơn vị!");
            txtFolderPath.focus();
            return;
        }

        const existingMaster = txtExistingMaster ? txtExistingMaster.value.trim() : "";
        const days = parseInt(spnDays.value, 10) || 16;

        btnRunStep1.disabled = true;
        loadingStep1.style.display = "flex";
        resultStep1.style.display = "none";

        try {
            const res = await fetch("/api/step1", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    folder_path: folderPath,
                    file_paths: selectedUnitFiles,
                    check_days: days,
                    existing_master: existingMaster
                })
            });

            const data = await res.json();

            if (res.ok && data.success) {
                currentWorkingDir = data.output_dir;

                // Hiển thị kết quả thống kê
                resTotalFiles.textContent = data.total_files;
                resTotalCases.textContent = Number(data.total_cases).toLocaleString("vi-VN");
                resFilteredCases.textContent = Number(data.cases_filtered).toLocaleString("vi-VN");
                pathFileTong.textContent = data.file_tong;
                pathFileKt.textContent = data.file_kt;
                timeStep1.textContent = new Date().toLocaleTimeString("vi-VN");

                if (data.mode === "append") {
                    boxAppendStats.style.display = "flex";
                    resExistingCases.textContent = Number(data.total_existing).toLocaleString("vi-VN");
                    resNewAdded.textContent = "+" + Number(data.new_added).toLocaleString("vi-VN");
                    resSkippedDupes.textContent = Number(data.skipped_duplicates).toLocaleString("vi-VN");
                    badgeStep1.textContent = "🔄 NẠP NỐI TIẾP THÀNH CÔNG";
                    lblTotalCases.textContent = "Tổng vụ trong File Tổng mới";

                    showToast(
                        "success",
                        "Nạp nối tiếp thành công!",
                        `Đã bảo lưu ${data.total_existing} vụ cũ, cập nhật ${data.updated_cases || data.skipped_duplicates} vụ và nạp thêm ${data.new_added} vụ mới vào File Tổng. Lọc ${data.cases_filtered} vụ án từ ${days} ngày trở lên!`
                    );
                } else {
                    boxAppendStats.style.display = "none";
                    badgeStep1.textContent = "✅ TỔNG HỢP THÀNH CÔNG";
                    lblTotalCases.textContent = "Tổng số vụ án gộp";

                    showToast(
                        "success",
                        "Tổng hợp hoàn tất!",
                        `Đã gộp ${data.total_files} file (${data.total_cases} vụ) và tách ${data.cases_filtered} vụ án từ ${days} ngày trở lên!`
                    );
                }

                resultStep1.style.display = "flex";

                // Tự động điền dữ liệu sang Bước 2
                txtFileTong.value = data.file_tong;
                txtFileKt.value = data.file_kt;
                stepIndicator2.classList.add("active");
            } else {
                showToast("error", "Không thể tổng hợp", data.message || data.error || "Có lỗi xảy ra khi xử lý file.");
            }
        } catch (err) {
            showToast("error", "Lỗi kết nối", "Không thể gửi yêu cầu xử lý đến Backend.");
        } finally {
            btnRunStep1.disabled = false;
            loadingStep1.style.display = "none";
        }
    });

    // 6. NÚT 2: Cập nhật kết quả và xuất Báo cáo Infographic
    btnRunStep2.addEventListener("click", async () => {
        const fileTong = txtFileTong.value.trim();
        const fileKt = txtFileKt.value.trim();

        if (!fileTong || !fileKt) {
            showToast("warning", "Thiếu file", "Vui lòng chọn đầy đủ cả 2 file:\n1. File Tổng hợp\n2. File Danh sách kiểm tra đã ghi chú!");
            return;
        }

        btnRunStep2.disabled = true;
        loadingStep2.style.display = "flex";
        resultStep2.style.display = "none";

        try {
            const res = await fetch("/api/step2", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ file_tong: fileTong, file_kt: fileKt })
            });

            const data = await res.json();

            if (res.ok && data.success) {
                currentWorkingDir = data.output_dir;

                // Hiển thị 4 thẻ KPI
                kpiTotal.textContent = Number(data.total_cases).toLocaleString("vi-VN");
                kpiKt.textContent = Number(data.total_checked).toLocaleString("vi-VN");
                kpiErrors.textContent = Number(data.total_errors).toLocaleString("vi-VN");
                kpiRate.textContent = data.error_rate;

                pathFileMaster.textContent = data.file_master;
                pathFileReport.textContent = data.file_report;
                timeStep2.textContent = new Date().toLocaleTimeString("vi-VN");

                // Hiển thị danh sách xếp hạng đơn vị vi phạm (nếu có)
                if (data.top_units && data.top_units.length > 0) {
                    unitTagsContainer.innerHTML = "";
                    data.top_units.forEach(u => {
                        const pill = document.createElement("span");
                        pill.className = "unit-pill";
                        pill.textContent = `${u["ĐƠN VỊ"]}: ${u["SỐ VỤ LỖI"]} vụ`;
                        unitTagsContainer.appendChild(pill);
                    });
                    unitRankingSection.style.display = "block";
                } else {
                    unitRankingSection.style.display = "none";
                }

                resultStep2.style.display = "flex";

                showToast(
                    "success",
                    "Xuất báo cáo thành công!",
                    `Đã tạo Báo cáo Infographic với ${data.total_errors} vụ phát hiện vi phạm/chậm tiến độ.`
                );
            } else {
                showToast("error", "Lỗi cập nhật", data.message || data.error || "Không thể cập nhật hoặc xuất báo cáo.");
            }
        } catch (err) {
            showToast("error", "Lỗi kết nối", "Không thể gửi yêu cầu xử lý đến Backend.");
        } finally {
            btnRunStep2.disabled = false;
            loadingStep2.style.display = "none";
        }
    });

    // 7. Mở thư mục kết quả trong Explorer
    async function openWorkingFolder() {
        const target = currentWorkingDir || txtFolderPath.value.trim() || txtFileTong.value.trim();
        if (!target) {
            showToast("warning", "Chưa có thư mục", "Vui lòng thực hiện Bước 1 hoặc chọn thư mục trước khi mở!");
            return;
        }

        try {
            const res = await fetch("/api/open-folder", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ path: target })
            });
            const data = await res.json();
            if (data.success) {
                showToast("success", "Đã mở thư mục", "Đã mở thư mục kết quả trong Windows Explorer.");
            } else {
                showToast("error", "Không thể mở", data.error || "Thư mục không tồn tại.");
            }
        } catch (err) {
            showToast("error", "Lỗi kết nối", "Không thể yêu cầu mở thư mục.");
        }
    }

    btnQuickOpenFolder.addEventListener("click", openWorkingFolder);
    btnFooterOpenFolder.addEventListener("click", openWorkingFolder);
});
