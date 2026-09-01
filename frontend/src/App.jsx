import { useEffect, useState } from "react";
import "./App.css";

// ------------------------------------------------------------
// АДРЕС BACKEND PRICE CONTROL V3
//
// React не подключается к PostgreSQL напрямую.
// Все операции с базой выполняются через Express API.
// ------------------------------------------------------------
const API_BASE_URL = "http://localhost:3001";

// ------------------------------------------------------------
// ПУСТАЯ ФОРМА МАТЕРИАЛА
//
// Используется:
// 1. при создании нового материала;
// 2. после очистки формы;
// 3. после завершения редактирования.
// ------------------------------------------------------------
function getEmptyMaterialForm() {
  return {
    resourceCode: "",
    officialName: "",
    baseQuantity: "1",
    baseUnit: "",
    verified: false,
    notes: "",
  };
}

function App() {
  // ------------------------------------------------------------
  // ОСНОВНЫЕ ДАННЫЕ СТРАНИЦЫ
  //
  // apiStatus — работает ли Express.
  // dbStatus  — есть ли подключение к PostgreSQL.
  // materials — список материалов из базы.
  // loading   — выполняется первоначальная загрузка.
  // error     — сообщение об ошибке.
  // ------------------------------------------------------------
  const [apiStatus, setApiStatus] = useState(null);
  const [dbStatus, setDbStatus] = useState(null);
  const [materials, setMaterials] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  // ------------------------------------------------------------
  // РЕДАКТИРОВАНИЕ СУЩЕСТВУЮЩЕГО МАТЕРИАЛА
  //
  // editingId — ID строки, которую пользователь редактирует.
  // editForm  — временно изменённые значения.
  // saving    — выполняется ли сейчас сохранение.
  // ------------------------------------------------------------
  const [editingId, setEditingId] = useState(null);

  const [editForm, setEditForm] = useState(getEmptyMaterialForm());

  const [saving, setSaving] = useState(false);

  // ------------------------------------------------------------
  // ДОБАВЛЕНИЕ НОВОГО МАТЕРИАЛА
  //
  // createForm — значения новой строки.
  // creating   — выполняется ли сейчас запись в PostgreSQL.
  // ------------------------------------------------------------
  const [createForm, setCreateForm] = useState(getEmptyMaterialForm());

  const [creating, setCreating] = useState(false);

  // ------------------------------------------------------------
  // УДАЛЕНИЕ МАТЕРИАЛА
  //
  // deletingId содержит ID материала,
  // который в данный момент удаляется.
  //
  // Это позволяет временно отключить кнопки именно этой строки
  // и показать пользователю текст "Удаление...".
  // ------------------------------------------------------------
  const [deletingId, setDeletingId] = useState(null);

  // ------------------------------------------------------------
  // ЗАГРУЗКА МАТЕРИАЛОВ ИЗ POSTGRESQL
  //
  // PostgreSQL
  //      ↓
  // Express
  //      ↓
  // GET /api/materials
  //      ↓
  // React
  //
  // После создания, изменения или удаления материала
  // эту функцию вызываем повторно.
  // ------------------------------------------------------------
  async function loadMaterials() {
    const response = await fetch(`${API_BASE_URL}/api/materials`);

    const data = await response.json();

    if (!response.ok || !data.ok) {
      throw new Error(data.message || "Не удалось получить материалы");
    }

    setMaterials(data.materials);
  }

  // ------------------------------------------------------------
  // ПЕРВОНАЧАЛЬНАЯ ЗАГРУЗКА СТРАНИЦЫ
  //
  // Проверяем:
  // 1. Express;
  // 2. PostgreSQL;
  // 3. список материалов.
  // ------------------------------------------------------------
  useEffect(() => {
    async function loadData() {
      try {
        setLoading(true);
        setError("");

        // Проверяем работу Express API.
        const apiResponse = await fetch(`${API_BASE_URL}/api/status`);

        const apiData = await apiResponse.json();

        setApiStatus(apiData);

        // Проверяем подключение Express к PostgreSQL.
        const dbResponse = await fetch(`${API_BASE_URL}/api/db/status`);

        const dbData = await dbResponse.json();

        setDbStatus(dbData);

        // Получаем материалы из PostgreSQL.
        await loadMaterials();
      } catch (loadError) {
        console.error("Ошибка загрузки данных:", loadError);

        setError(loadError.message);
      } finally {
        setLoading(false);
      }
    }

    loadData();
  }, []);

  // ------------------------------------------------------------
  // НАЧАЛО РЕДАКТИРОВАНИЯ МАТЕРИАЛА
  //
  // Текущие значения строки копируются в editForm.
  // PostgreSQL на этом этапе ещё не изменяется.
  // ------------------------------------------------------------
  function startEditing(material) {
    setEditingId(material.id);

    setEditForm({
      resourceCode: material.resource_code || "",
      officialName: material.official_name || "",
      baseQuantity: material.base_quantity || "",
      baseUnit: material.base_unit || "",
      verified: material.verified,
      notes: material.notes || "",
    });
  }

  // ------------------------------------------------------------
  // ОТМЕНА РЕДАКТИРОВАНИЯ
  //
  // Несохранённые изменения удаляются.
  // База данных при этом не меняется.
  // ------------------------------------------------------------
  function cancelEditing() {
    setEditingId(null);
    setEditForm(getEmptyMaterialForm());
  }

  // ------------------------------------------------------------
  // ИЗМЕНЕНИЕ ПОЛЕЙ РЕДАКТИРОВАНИЯ
  // ------------------------------------------------------------
  function handleEditChange(event) {
    const { name, value, type, checked } = event.target;

    setEditForm((currentForm) => ({
      ...currentForm,

      [name]: type === "checkbox" ? checked : value,
    }));
  }

  // ------------------------------------------------------------
  // СОХРАНЕНИЕ ИЗМЕНЁННОГО МАТЕРИАЛА
  //
  // React
  //   ↓
  // PATCH /api/materials/:id
  //   ↓
  // Express
  //   ↓
  // PostgreSQL
  // ------------------------------------------------------------
  async function saveMaterial(id) {
    try {
      setSaving(true);
      setError("");

      // Наименование материала обязательно.
      if (!editForm.officialName.trim()) {
        throw new Error("Наименование материала не может быть пустым");
      }

      // Преобразуем количество в число.
      const baseQuantity = Number(editForm.baseQuantity);

      if (!Number.isFinite(baseQuantity)) {
        throw new Error("Количество должно быть числом");
      }

      // Передаём изменения в backend.
      const response = await fetch(`${API_BASE_URL}/api/materials/${id}`, {
        method: "PATCH",

        headers: {
          "Content-Type": "application/json; charset=utf-8",
        },

        body: JSON.stringify({
          resourceCode: editForm.resourceCode.trim() || null,

          officialName: editForm.officialName.trim(),

          baseQuantity,

          baseUnit: editForm.baseUnit.trim() || null,

          verified: editForm.verified,

          notes: editForm.notes.trim() || null,
        }),
      });

      const data = await response.json();

      if (!response.ok || !data.ok) {
        throw new Error(data.message || "Не удалось сохранить материал");
      }

      // После сохранения заново читаем данные из PostgreSQL.
      await loadMaterials();

      // Закрываем режим редактирования.
      setEditingId(null);
      setEditForm(getEmptyMaterialForm());
    } catch (saveError) {
      console.error("Ошибка сохранения материала:", saveError);

      setError(saveError.message);
    } finally {
      setSaving(false);
    }
  }

  // ------------------------------------------------------------
  // ИЗМЕНЕНИЕ ПОЛЕЙ НОВОГО МАТЕРИАЛА
  // ------------------------------------------------------------
  function handleCreateChange(event) {
    const { name, value, type, checked } = event.target;

    setCreateForm((currentForm) => ({
      ...currentForm,

      [name]: type === "checkbox" ? checked : value,
    }));
  }

  // ------------------------------------------------------------
  // СОЗДАНИЕ НОВОГО МАТЕРИАЛА
  //
  // React
  //   ↓
  // POST /api/materials
  //   ↓
  // Express
  //   ↓
  // INSERT INTO materials
  //   ↓
  // PostgreSQL
  // ------------------------------------------------------------
  async function createMaterial(event) {
    event.preventDefault();

    try {
      setCreating(true);
      setError("");

      // Наименование обязательно.
      if (!createForm.officialName.trim()) {
        throw new Error("Укажите наименование материала");
      }

      // Количество переводим из текста в число.
      const baseQuantity = Number(createForm.baseQuantity);

      if (!Number.isFinite(baseQuantity)) {
        throw new Error("Количество должно быть числом");
      }

      // Отправляем новую запись в backend.
      const response = await fetch(`${API_BASE_URL}/api/materials`, {
        method: "POST",

        headers: {
          "Content-Type": "application/json; charset=utf-8",
        },

        body: JSON.stringify({
          resourceCode: createForm.resourceCode.trim() || null,

          officialName: createForm.officialName.trim(),

          baseQuantity,

          baseUnit: createForm.baseUnit.trim() || null,

          verified: createForm.verified,

          notes: createForm.notes.trim() || null,
        }),
      });

      const data = await response.json();

      if (!response.ok || !data.ok) {
        throw new Error(data.message || "Не удалось создать материал");
      }

      // После INSERT перечитываем PostgreSQL.
      await loadMaterials();

      // Очищаем форму.
      setCreateForm(getEmptyMaterialForm());
    } catch (createError) {
      console.error("Ошибка создания материала:", createError);

      setError(createError.message);
    } finally {
      setCreating(false);
    }
  }

  // ------------------------------------------------------------
  // УДАЛЕНИЕ МАТЕРИАЛА
  //
  // Это важная операция, потому что запись физически
  // удаляется из PostgreSQL.
  //
  // Поэтому сначала обязательно спрашиваем подтверждение.
  //
  // Если пользователь нажимает "Отмена",
  // никакого DELETE-запроса вообще не выполняется.
  // ------------------------------------------------------------
  async function deleteMaterial(material) {
    // Показываем название материала,
    // чтобы пользователь точно видел, что именно удаляет.
    const confirmed = window.confirm(
      `Удалить материал "${material.official_name}"?\n\n` +
        "Это действие удалит запись из PostgreSQL.",
    );

    // Пользователь нажал "Отмена".
    // Сразу прекращаем выполнение функции.
    if (!confirmed) {
      return;
    }

    try {
      setDeletingId(material.id);
      setError("");

      // --------------------------------------------------------
      // ОТПРАВЛЯЕМ DELETE В BACKEND
      //
      // Например:
      //
      // DELETE /api/materials/2
      //
      // Backend затем выполняет DELETE в PostgreSQL.
      // --------------------------------------------------------
      const response = await fetch(
        `${API_BASE_URL}/api/materials/${material.id}`,
        {
          method: "DELETE",
        },
      );

      const data = await response.json();

      // Если backend сообщил об ошибке,
      // показываем её пользователю.
      if (!response.ok || !data.ok) {
        throw new Error(data.message || "Не удалось удалить материал");
      }

      // Если удалялась строка, которая была открыта
      // в режиме редактирования, закрываем этот режим.
      if (editingId === material.id) {
        setEditingId(null);
        setEditForm(getEmptyMaterialForm());
      }

      // После DELETE снова читаем материалы из PostgreSQL.
      //
      // Благодаря этому таблица React отражает
      // реальное состояние базы.
      await loadMaterials();
    } catch (deleteError) {
      console.error("Ошибка удаления материала:", deleteError);

      setError(deleteError.message);
    } finally {
      setDeletingId(null);
    }
  }

  // ------------------------------------------------------------
  // ОСНОВНОЙ ИНТЕРФЕЙС PRICE CONTROL V3
  // ------------------------------------------------------------
  return (
    <div className="app">
      <h1>Price Control V3</h1>

      {/* --------------------------------------------------------
          ТЕХНИЧЕСКОЕ СОСТОЯНИЕ СИСТЕМЫ
      --------------------------------------------------------- */}
      <section>
        <h2>Состояние системы</h2>

        <p>API: {apiStatus?.message || "Проверка..."}</p>

        <p>
          PostgreSQL:{" "}
          {dbStatus?.ok
            ? `подключен: ${dbStatus.database}, порт ${dbStatus.port}`
            : "Проверка..."}
        </p>
      </section>

      {/* --------------------------------------------------------
          ОШИБКИ
      --------------------------------------------------------- */}
      {error && <p>Ошибка: {error}</p>}

      {/* --------------------------------------------------------
          ДОБАВЛЕНИЕ НОВОГО МАТЕРИАЛА
      --------------------------------------------------------- */}
      <section>
        <h2>Добавить материал</h2>

        <form onSubmit={createMaterial}>
          <div>
            <label>
              Код ресурса:{" "}
              <input
                name="resourceCode"
                value={createForm.resourceCode}
                onChange={handleCreateChange}
              />
            </label>
          </div>

          <div>
            <label>
              Наименование:{" "}
              <input
                name="officialName"
                value={createForm.officialName}
                onChange={handleCreateChange}
                required
              />
            </label>
          </div>

          <div>
            <label>
              Количество:{" "}
              <input
                name="baseQuantity"
                type="number"
                step="any"
                value={createForm.baseQuantity}
                onChange={handleCreateChange}
              />
            </label>
          </div>

          <div>
            <label>
              Ед. изм.:{" "}
              <input
                name="baseUnit"
                value={createForm.baseUnit}
                onChange={handleCreateChange}
              />
            </label>
          </div>

          <div>
            <label>
              <input
                name="verified"
                type="checkbox"
                checked={createForm.verified}
                onChange={handleCreateChange}
              />{" "}
              Материал проверен
            </label>
          </div>

          <div>
            <label>
              Примечание:{" "}
              <input
                name="notes"
                value={createForm.notes}
                onChange={handleCreateChange}
              />
            </label>
          </div>

          <button type="submit" disabled={creating}>
            {creating ? "Добавление..." : "Добавить материал"}
          </button>
        </form>
      </section>

      {/* --------------------------------------------------------
          ТАБЛИЦА МАТЕРИАЛОВ
      --------------------------------------------------------- */}
      <section>
        <h2>Материалы</h2>

        {loading && <p>Загрузка материалов...</p>}

        {!loading && materials.length === 0 && (
          <p>В базе пока нет материалов.</p>
        )}

        {!loading && materials.length > 0 && (
          <table>
            <thead>
              <tr>
                <th>ID</th>
                <th>Код ресурса</th>
                <th>Наименование</th>
                <th>Количество</th>
                <th>Ед. изм.</th>
                <th>Проверен</th>
                <th>Примечание</th>
                <th>Действия</th>
              </tr>
            </thead>

            <tbody>
              {materials.map((material) => {
                // Определяем состояние конкретной строки.
                const isEditing = editingId === material.id;

                const isDeleting = deletingId === material.id;

                return (
                  <tr key={material.id}>
                    <td>{material.id}</td>

                    {/* Код ресурса */}
                    <td>
                      {isEditing ? (
                        <input
                          name="resourceCode"
                          value={editForm.resourceCode}
                          onChange={handleEditChange}
                        />
                      ) : (
                        material.resource_code || ""
                      )}
                    </td>

                    {/* Наименование */}
                    <td>
                      {isEditing ? (
                        <input
                          name="officialName"
                          value={editForm.officialName}
                          onChange={handleEditChange}
                        />
                      ) : (
                        material.official_name
                      )}
                    </td>

                    {/* Количество */}
                    <td>
                      {isEditing ? (
                        <input
                          name="baseQuantity"
                          type="number"
                          step="any"
                          value={editForm.baseQuantity}
                          onChange={handleEditChange}
                        />
                      ) : (
                        material.base_quantity
                      )}
                    </td>

                    {/* Единица измерения */}
                    <td>
                      {isEditing ? (
                        <input
                          name="baseUnit"
                          value={editForm.baseUnit}
                          onChange={handleEditChange}
                        />
                      ) : (
                        material.base_unit || ""
                      )}
                    </td>

                    {/* Материал проверен */}
                    <td>
                      {isEditing ? (
                        <input
                          name="verified"
                          type="checkbox"
                          checked={editForm.verified}
                          onChange={handleEditChange}
                        />
                      ) : material.verified ? (
                        "Да"
                      ) : (
                        "Нет"
                      )}
                    </td>

                    {/* Примечание */}
                    <td>
                      {isEditing ? (
                        <input
                          name="notes"
                          value={editForm.notes}
                          onChange={handleEditChange}
                        />
                      ) : (
                        material.notes || ""
                      )}
                    </td>

                    {/* ------------------------------------------------
                        ДЕЙСТВИЯ С МАТЕРИАЛОМ

                        Здесь находятся:
                        - редактирование;
                        - сохранение;
                        - отмена;
                        - удаление.
                    ------------------------------------------------- */}
                    <td>
                      {isEditing ? (
                        <>
                          <button
                            type="button"
                            onClick={() => saveMaterial(material.id)}
                            disabled={saving || isDeleting}
                          >
                            {saving ? "Сохранение..." : "Сохранить"}
                          </button>

                          <button
                            type="button"
                            onClick={cancelEditing}
                            disabled={saving || isDeleting}
                          >
                            Отмена
                          </button>
                        </>
                      ) : (
                        <>
                          <button
                            type="button"
                            onClick={() => startEditing(material)}
                            disabled={isDeleting}
                          >
                            Изменить
                          </button>

                          <button
                            type="button"
                            onClick={() => deleteMaterial(material)}
                            disabled={isDeleting}
                          >
                            {isDeleting ? "Удаление..." : "Удалить"}
                          </button>
                        </>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </section>
    </div>
  );
}

export default App;
