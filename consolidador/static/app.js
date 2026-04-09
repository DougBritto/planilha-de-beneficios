document.addEventListener("DOMContentLoaded", () => {
  const fileInputs = document.querySelectorAll("[data-file-input]");
  const selectableTables = document.querySelectorAll("[data-selectable-table]");

  fileInputs.forEach((fileInput) => {
    const key = fileInput.getAttribute("data-file-input");
    const fileList = document.querySelector(`[data-file-list="${key}"]`);
    if (!fileList) {
      return;
    }

    const emptyLabel =
      key === "saude"
        ? "Nenhum arquivo de Saude selecionado."
        : "Nenhum arquivo de Odonto selecionado.";

    const renderSelectedFiles = () => {
      const files = Array.from(fileInput.files || []);
      if (!files.length) {
        fileList.textContent = emptyLabel;
        return;
      }

      const labels = files.map((file) => `${file.name} (${Math.round(file.size / 1024)} KB)`);
      fileList.textContent = labels.join(" | ");
    };

    fileInput.addEventListener("change", renderSelectedFiles);
  });

  selectableTables.forEach((tableWrapper) => {
    const selectAll = tableWrapper.querySelector("[data-select-all]");
    const items = Array.from(tableWrapper.querySelectorAll("[data-select-item]"));
    if (!selectAll || !items.length) {
      return;
    }

    const syncSelectAllState = () => {
      const checkedItems = items.filter((item) => item.checked).length;
      selectAll.checked = checkedItems > 0 && checkedItems === items.length;
      selectAll.indeterminate = checkedItems > 0 && checkedItems < items.length;
    };

    selectAll.addEventListener("change", () => {
      items.forEach((item) => {
        item.checked = selectAll.checked;
      });
      syncSelectAllState();
    });

    items.forEach((item) => {
      item.addEventListener("change", syncSelectAllState);
    });

    syncSelectAllState();
  });
});
