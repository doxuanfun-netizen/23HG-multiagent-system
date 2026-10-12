# DANH MỤC 135 LỆNH ĐIỀU KHIỂN CAD TỰ ĐỘNG BẰNG TIẾNG VIỆT
**23HG CAD AUTOMATION MCP PRO (v2.0.0)**  
*Tác giả: Nguyễn Bảo Tú (@baotuhg) | 23HG SYSTEM*

> [!IMPORTANT]
> **THÔNG BÁO KIẾN TRÚC & PHÂN TÁCH CÔNG NGHỆ (v2.0.0):**
> Danh mục 135 lệnh dưới đây là bảng đặc tả của plugin .NET DLL cũ (Legacy Reference). Kể từ phiên bản v2.0.0, toàn bộ kiến trúc lõi của hệ thống đã được chuyển đổi hoàn toàn sang **Native COM Interop & QS Takeoff Engine**:
> 1. **Zero-Plugin / Zero-DLL**: Không cần gõ lệnh `NETLOAD`, không lo xung đột phiên bản AutoCAD (2020 - 2026+).
> 2. **Động cơ Bóc tách Khối lượng thực tế (QS Takeoff COM Engine)**: 4 công cụ bóc tách cốt lõi (`cad_get_units`, `cad_quantity_by_layer`, `cad_count_blocks`, `cad_export_boq`) thực thi trực tiếp qua ActiveX Automation COM, hỗ trợ bóc tách hình học, chống cộng trùng diện tích, giải mã Dynamic Block và xuất Excel 100% công thức sống.
> 3. Bảng 135 lệnh này được lưu giữ để tham chiếu cấu trúc tham số; khi vận hành thực tế qua MCP, các AI Client sẽ tương tác trực tiếp qua Native COM Tools và SendCommand an toàn (có Whitelist và Security Guard bảo vệ).

---

## 1. Vẽ hình cơ bản (10 lệnh)

| Tên Lệnh API | Tên CAD / NetLoad | Mô Tả Chức Năng | Tham Số Chính |
| :--- | :--- | :--- | :--- |
| api_create_arc |  | Tạo cung tròn Arc bằng tâm, bán kính, góc bắt đầu và góc kết thúc | center, radius, startAngle, endAngle, layerName, colorIndex |
| api_create_dimension |  | Tạo đối tượng kích thước Dim (thẳng, gióng nghiêng, góc...) | dimType, point1, point2, point3, dimLinePoint, dimensionText, rotation, entityHandle, layerName |
| api_create_hatch |  | Tạo mẫu mặt cắt Hatch dựa trên đường bao kín hoặc danh sách điểm | boundaryHandles, points, patternName, scale, patternAngle, hatchStyle, associative, layerName, colorIndex |
| api_create_leader |  | Tạo đường dẫn Leader có kèm theo chữ chú thích | points, annotation, annotationHeight, hasArrowhead, splineFit, layerName, colorIndex |
| api_create_mtext |  | Tạo văn bản nhiều dòng MText, hỗ trợ xuống dòng và định dạng | position, text, width, height, rotation, attachment, textStyle, layerName, colorIndex |
| api_create_polyline |  | Tạo đường đa tuyến Pline, hỗ trợ đóng kín và tạo cung cong | points, closed, layerName, colorIndex, constantWidth |
| api_create_text |  | Tạo văn bản đơn lẻ một dòng DText | position, text, height, rotation, widthFactor, textStyle, horizontalMode, layerName, colorIndex |
| api_draw_circle |  | Vẽ đường tròn Circle theo tâm và bán kính | center, radius, layerName, colorIndex |
| api_draw_line |  | Vẽ đường thẳng Line từ điểm bắt đầu đến điểm kết thúc | startPoint, endPoint, layerName, colorIndex |
| api_draw_rectangle |  | Vẽ hình chữ nhật bằng 2 điểm góc đối diện | corner1, corner2, layerName, colorIndex |

## 2. Truy vấn đối tượng (16 lệnh)

| Tên Lệnh API | Tên CAD / NetLoad | Mô Tả Chức Năng | Tham Số Chính |
| :--- | :--- | :--- | :--- |
| api_batch_get_entity_info |  | Lấy thông tin chi tiết của nhiều thực thể theo danh sách Handle | entityHandles, layerName, entityType |
| api_get_block_attributes |  | Lấy danh sách các cặp Tag/Value thuộc tính Attribute của Block | entityHandle |
| api_get_block_definition |  | Lấy thông tin chi tiết định nghĩa cấu trúc của Block trong bản vẽ | blockName |
| api_get_blocks_by_name |  | Tìm kiếm tất cả các tham chiếu Block theo tên Block chỉ định | blockName, matchMode, currentSpaceOnly |
| api_get_entity_info |  | Lấy toàn bộ thuộc tính chi tiết của thực thể theo Handle | entityHandle |
| api_get_unique_attribute_values |  | Lấy các giá trị độc nhất của một thuộc tính Attribute | attributeName, blockName, prefix, currentSpaceOnly, maxBlocks, includeEmpty |
| api_get_xdata |  | Đọc dữ liệu mở rộng XData gắn trên thực thể theo AppName | entityHandle, entityHandles, appName |
| api_query_all_block_names |  | Lấy danh sách toàn bộ các Block có trong bản vẽ | includeAnonymous, includeModelSpace, includeXrefs, pattern, includeReferenceCount |
| api_query_by_attribute |  | Tìm các Block có giá trị thuộc tính Attribute khớp với điều kiện | attributeName, attributeValue, matchMode, blockName, currentSpaceOnly, maxCount |
| api_query_by_layer |  | Tìm kiếm tất cả thực thể nằm trên Layer được chỉ định | layerName, entityTypes, currentSpaceOnly, maxCount |
| api_query_by_polygon |  | Tìm các thực thể nằm trong hoặc cắt qua vùng đa giác khép kín | points, crossing, entityTypes, layerName, maxCount |
| api_query_by_type |  | Tìm kiếm đối tượng theo loại (Line, Polyline, Circle, Text...) | entityType, layerName, currentSpaceOnly, maxCount |
| api_query_by_window |  | Tìm các thực thể nằm trong hoặc cắt qua vùng chữ nhật chọn | corner1, corner2, crossing, entityTypes, layerName, maxCount |
| api_query_connected_entities |  | Tìm các đối tượng có điểm đầu/cuối nối tiếp với thực thể gốc | entityHandle, tolerance, entityTypes, maxDepth |
| api_query_intersecting |  | Tìm các đối tượng cắt qua hoặc giao cắt với thực thể chỉ định | entityHandle, entityTypes, layerName, maxCount |
| api_query_nearby_entities |  | Tìm các thực thể gần một tọa độ điểm trong bán kính xác định | point, radius, entityTypes, layerName, maxCount, sortByDistance |

## 3. Thao tác Block (9 lệnh)

| Tên Lệnh API | Tên CAD / NetLoad | Mô Tả Chức Năng | Tham Số Chính |
| :--- | :--- | :--- | :--- |
| api_create_block_definition |  | Tạo một định nghĩa Block mới từ danh sách thực thể | blockName, basePoint, entityHandles, deleteOriginal, attributeDefinitions |
| api_explode_block |  | Phá vỡ Block thành các thực thể hình học độc lập | entityHandle, deleteOriginal, convertAttributesToText |
| api_import_block_from_external |  | Nạp định nghĩa Block từ file DWG bên ngoài vào bản vẽ hiện hành | blockName, blockNames, sourceDwgPath |
| api_insert_block |  | Chèn tham chiếu Block vào bản vẽ tại tọa độ điểm, tỷ lệ và góc xoay | blockName, insertionPoint, xScale, yScale, zScale, rotation, layerName, attributes |
| api_modify_block_base_point |  | Định nghĩa lại tọa độ điểm gốc (Base Point) của Block | entityHandle, newBasePoint |
| api_modify_block_definition |  | Chỉnh sửa cấu trúc hình học bên trong định nghĩa Block | blockName, newBasePoint, addEntityHandles, removeInternalIndexes |
| api_rename_block |  | Đổi tên định nghĩa Block trong bản vẽ | oldBlockName, newBlockName, prefix, suffix, pattern |
| api_set_block_attributes |  | Gán giá trị thuộc tính Attribute hoặc Dynamic Block | entityHandle, attributes, dynamicAttributes |
| api_sync_block_references |  | Đồng bộ hóa các thay đổi thuộc tính tới toàn bộ Block tham chiếu | blockName, preserveAttributeValues, currentSpaceOnly |

## 4. Thao tác hàng loạt (13 lệnh)

| Tên Lệnh API | Tên CAD / NetLoad | Mô Tả Chức Năng | Tham Số Chính |
| :--- | :--- | :--- | :--- |
| api_batch_bring_to_behind |  | Thay đổi thứ tự hiển thị đưa đối tượng ra sau | entityHandles |
| api_batch_bring_to_front |  | Thay đổi thứ tự hiển thị đưa đối tượng lên trước | entityHandles |
| api_batch_convert_lines_to_polylines |  | Nối các đoạn Line nối tiếp thành đường đa tuyến Polyline khép kín | entityHandles, deleteOriginal |
| api_batch_copy |  | Sao chép cùng lúc nhiều đối tượng theo vector độ dời | entityHandles, displacement |
| api_batch_create_texts |  | Tạo cùng lúc nhiều đối tượng Text từ mảng dữ liệu hoặc file JSON | texts, jsonFilePath, layerName, createLayerIfNotExists |
| api_batch_get_block_attributes |  | Lấy toàn bộ giá trị Attribute của danh sách Block | entityHandles, attributeTags |
| api_batch_insert_blocks |  | Chèn cùng lúc nhiều Block tại danh sách các tọa độ điểm | insertions, blockName, insertionPoints, scales, rotations, layerName, createLayerIfNotExists |
| api_batch_move |  | Di chuyển cùng lúc nhiều đối tượng theo vector độ dời | entityHandles, displacement |
| api_batch_replace_blocks |  | Thay thế một loại Block bằng loại Block khác trên toàn bản vẽ | oldBlockName, newBlockName, rotationAngle, scale, entityHandles, filterAttrName, filterAttrPattern |
| api_batch_set_attributes |  | Cập nhật thuộc tính Attribute cho danh sách nhiều Block | entityHandles, attributes |
| api_batch_set_layer |  | Chuyển cùng lúc nhiều đối tượng sang một Layer mới | entityHandles, layerName |
| api_batch_transform |  | Áp dụng ma trận biến đổi tọa độ cho nhiều thực thể | entityHandles, transformType, transformMatrix, basePoint, displacement, angle, scaleFactor, matrix |
| api_move_glyph_strokes_to_layer |  | Tách các nét vector glyph vẽ chữ PDF sang layer riêng | sourceTextLayer, targetLayer, excludeLayers, padding, maxStrokeSize, createLayerIfNotExists |

## 5. Tính toán không gian hình học (8 lệnh)

| Tên Lệnh API | Tên CAD / NetLoad | Mô Tả Chức Năng | Tham Số Chính |
| :--- | :--- | :--- | :--- |
| api_calc_angle |  | Tính góc hợp bởi 2 vector hoặc 3 điểm tọa độ | point1, point2, point3 |
| api_calc_area |  | Tính diện tích hình học khép kín theo thuật toán Shoelace Gauss | entityHandle, points |
| api_calc_bounding_box |  | Tính tọa độ góc nhỏ nhất và lớn nhất bao quanh đối tượng | entityHandle, entityHandles |
| api_calc_centroid |  | Tính tọa độ tâm hình học trọng tâm của đa giác khép kín | entityHandle |
| api_calc_distance |  | Tính khoảng cách không gian giữa 2 điểm tọa độ | point1, point2, entityHandle1, entityHandle2 |
| api_calc_intersection |  | Tính tọa độ giao điểm giữa 2 thực thể hình học | handle1, handle2, extendMode |
| api_calc_nearest_point |  | Tìm điểm trên thực thể có khoảng cách ngắn nhất tới điểm chỉ định | entityHandle, point, extend |
| api_point_on_entity |  | Kiểm tra xem một điểm tọa độ có nằm trên hoặc thuộc đường biên thực thể không | point, entityHandle, tolerance |

## 6. Quản lý Layer (1 lệnh)

| Tên Lệnh API | Tên CAD / NetLoad | Mô Tả Chức Năng | Tham Số Chính |
| :--- | :--- | :--- | :--- |
| api_import_layer_from_external |  | Sao chép cấu hình Layer từ file DWG khác vào bản vẽ | layerName, layerNames, sourceDwgPath |

## 7. Tập hợp chọn (Selection Set) (8 lệnh)

| Tên Lệnh API | Tên CAD / NetLoad | Mô Tả Chức Năng | Tham Số Chính |
| :--- | :--- | :--- | :--- |
| api_export_selection_preview |  | Tạo ảnh preview vùng đối tượng được chọn | entityHandles, backgroundColor, padding, maxWidth, maxHeight |
| api_get_selection_set |  | Lấy danh sách các Handle của các đối tượng đang được người dùng chọn | includeDetails, promptSelect |
| api_select_all |  | Chọn toàn bộ các thực thể trong không gian vẽ hiện hành | entityType, layerName |
| api_select_by_filter |  | Chọn các thực thể thỏa mãn điều kiện loại đối tượng và Layer | entityTypes, layerNames, colorIndex, blockName |
| api_select_by_xdata |  | Chọn các thực thể có chứa dữ liệu mở rộng XData khớp điều kiện | appName, matchKey, matchValue, inverse, layerName, entityTypes, currentSpaceOnly, maxCount, includeDetails, highlight |
| api_select_fence |  | Chọn các đối tượng bị đường cắt Fence cắt qua | points, entityType |
| api_select_window |  | Chọn các đối tượng nằm trong hoặc giao cắt khung cửa sổ | point1, point2, crossingMode, entityType |
| api_window_select_with_filter |  | Quét chọn theo vùng chữ nhật có kèm lọc loại đối tượng | entityTypes, layerNames, colorIndex, blockName, promptMessage, useCrossingMode |

## 8. Sơ đồ công nghệ P&ID GsLc (36 lệnh)

| Tên Lệnh API | Tên CAD / NetLoad | Mô Tả Chức Năng | Tham Số Chính |
| :--- | :--- | :--- | :--- |
| api_gslc_auto_number |  | Công cụ đánh số tự động hỗ trợ đánh số thông minh cho đường ống, thiết bị, dụng cụ và van | dataType, numberMode, numberDirection, digits, prefix, startNumber, selectionMode, entityHandles, customBlockName, previewOnly, frameHandle, sortTolerance |
| api_gslc_brush_instrument_location |  | Quét vị trí của thiết bị bằng một cú nhấp chuột và liên kết vị trí tương đối: Thực hiện theo dõi cấu trúc liên kết kết nối/đường ống để liên kết thiết bị với thiết bị (số vị trí) hoặc đường ống (số đường ống), ghi các thuộc tính vị trí và dữ liệu vị trí tương đối và di chuyển nó đến lớp đông lạnh. Theo mặc định, toàn bộ hình ảnh được xử lý; nếu frameHandle được thông qua, chỉ phạm vi sẽ được xử lý. | frameHandle, onlyUnprocessed, setFreezeLayer, includeUnmatched, equipCandidates |
| api_gslc_brush_join_draw_arrow |  | Làm mới các thuộc tính hướng và đường ống liên quan cho các mũi tên sơ đồ | arrowHandle, pipeHandle, fromTo, drawNum |
| api_gslc_brush_pipe_from_to |  | Duyệt hướng của đường ống bằng một cú nhấp chuột và liên kết vị trí tương đối: đối với mỗi khối số đường ống, theo dõi thiết bị (số vị trí)/mã chung ở cả hai đầu dọc theo cấu trúc liên kết đường ống, xác định hướng/đích bằng mũi tên luồng, viết thuộc tính hướng/đích đường ống và vị trí tương đối của hai điểm và di chuyển nó đến lớp cố định. Theo mặc định, toàn bộ hình ảnh được xử lý; nếu frameHandle được thông qua, chỉ phạm vi sẽ được xử lý. | frameHandle, onlyUnprocessed, setFreezeLayer, includeUnmatched |
| api_gslc_clear_relative_position |  | Xóa dữ liệu vị trí tương đối cho một hoặc nhiều thực thể | entityHandles, clearLocationAttr |
| api_gslc_export_all_data |  | Xuất tất cả dữ liệu trong sơ đồ quy trình (thiết bị, dụng cụ, đường ống, van và phụ kiện) sang định dạng JArray phẳng | selectionMode, entityHandles, outputPath |
| api_gslc_export_frame_data |  | Xuất dữ liệu sơ đồ quy trình. Bạn có thể chỉ định phạm vi khung vẽ hoặc xuất toàn bộ bản vẽ | frameHandle, includeRelativePosition, includeAttributes |
| api_gslc_generate_equipment_list |  | Tạo dữ liệu danh sách thiết bị trong khung được chỉ định | frameHandle, groupByType, includeEmpty |
| api_gslc_insert_bottom_equip_tags |  | Tự động tạo khối thẻ thiết bị ở cuối tất cả các khung vẽ | frameHandles, zoomToFrames |
| api_gslc_insert_control_valve |  | Lắp các van điều khiển (van điều chỉnh, van chuyển mạch…) tại các vị trí được chỉ định | valveType, position, instrumentNum, location, scale, rotation |
| api_gslc_insert_equip_tag |  | Chèn nhãn thẻ thiết bị vào vị trí được chỉ định | position, tag, name, relatedEquipHandle, scale, rotation |
| api_gslc_insert_equipment |  | Chèn một khối thiết bị vào vị trí được chỉ định | equipType, position, tag, name, description, scale, rotation |
| api_gslc_insert_instrument |  | Chèn khối thiết bị vào vị trí đã chỉ định | instrumentType, position, instrumentNum, location, scale, rotation |
| api_gslc_insert_join_draw_arrow |  | Chèn một mũi tên kết nối vào vị trí đã chỉ định | direction, position, fromTo, pipeNum, drawNum, scale, rotation |
| api_gslc_insert_outer_pipe_arrow |  | Chèn mũi tên ống ngoài vào vị trí được chỉ định | arrowType, position, pipeNum, fromTo, scale, rotation |
| api_gslc_insert_pipe |  | Chèn khối nhãn ống vào vị trí đã chỉ định | position, pipeNum, from, to, pipeClass, pipeSize, pipeType, scale, rotation |
| api_gslc_insert_valve |  | Chèn một van thông thường vào vị trí được chỉ định | valveType, position, valveNum, location, scale, rotation |
| api_gslc_move_to_freeze_layer |  | Di chuyển các thực thể sang một lớp cố định (cho biết dữ liệu đã được xử lý) | entityHandles, freezeLayerType |
| api_gslc_query_all_entities_in_frame |  | Truy vấn tất cả các thực thể sơ đồ quy trình (thiết bị, dụng cụ, van, đường ống, mũi tên kết nối). Bạn có thể chỉ định phạm vi khung vẽ hoặc truy vấn toàn bộ bản vẽ | frameHandle, entityTypes |
| api_gslc_query_entities_by_relative_position |  | Truy vấn tất cả các đồng hồ hoặc đường ống liên quan đến thực thể được chỉ định thông qua vị trí tương đối | referenceHandle, frameHandle, entityType |
| api_gslc_query_equipment_in_frame |  | Truy vấn thông tin thiết bị. Bạn có thể chỉ định phạm vi khung vẽ hoặc truy vấn toàn bộ bản vẽ | frameHandle, equipType |
| api_gslc_query_frame_info |  | Truy vấn thông tin khung của sơ đồ quy trình. Trả về tất cả các khung khi frameHandle không được chỉ định | frameHandle, includeStats |
| api_gslc_query_instrument_in_frame |  | Truy vấn thông tin thiết bị (bao gồm cả van điều khiển). Bạn có thể chỉ định phạm vi khung vẽ hoặc truy vấn toàn bộ bản vẽ | frameHandle, instrumentType, hasLocation, hasRelativePosition |
| api_gslc_query_join_draw_arrow_in_frame |  | Truy vấn thông tin của mũi tên kết nối. Bạn có thể chỉ định phạm vi khung vẽ hoặc truy vấn toàn bộ bản vẽ | frameHandle, direction |
| api_gslc_query_pipe_in_frame |  | Truy vấn thông tin đường ống. Bạn có thể chỉ định phạm vi khung vẽ hoặc truy vấn toàn bộ bản vẽ | frameHandle, hasFromTo, hasRelativePosition |
| api_gslc_query_valve_in_frame |  | Truy vấn thông tin van (không bao gồm van điều khiển). Bạn có thể chỉ định phạm vi khung vẽ hoặc truy vấn toàn bộ bản vẽ | frameHandle, valveType |
| api_gslc_replace_instrument_block |  | Chuyển đổi mét số lượng lớn thành khối mét. Nhận dạng tự động: hình tròn + hình chữ nhật ngoại vi → khối thiết bị truyền từ xa, chỉ vòng tròn → khối thiết bị cục bộ | layerName, minRadius, maxRadius, targetLayerName, deleteOriginal, previewOnly, scale |
| api_gslc_restore_from_freeze_layer |  | Di chuyển các thực thể từ lớp cố định trở lại lớp ban đầu | entityHandles, targetLayer |
| api_gslc_set_instrument_location |  | Liên kết rõ ràng vị trí của nhạc cụ (nguyên thủy duy trì phối hợp linh hoạt): được cung cấp tay cầm nhạc cụ + tay cầm khối tham chiếu (khối số ống/khối thẻ thiết bị), đọc vị trí thuộc tính khối tham chiếu (có thể được ghi đè bằng vị trí), viết thuộc tính vị trí + vị trí tương đối một điểm + di chuyển lớp cố định, phù hợp với lệnh cọ vẽ. Hỗ trợ một mục duy nhất (instrumentHandle/referenceHandle) hoặc hàng loạt (items[]); bảng chỉ báo đôi và bộ phát chia sẻ cùng một tay cầm tham chiếu và có thể kế thừa cùng một vị trí. | instrumentHandle, referenceHandle, location, items, setFreezeLayer |
| api_gslc_set_pipe_from_to |  | Liên kết rõ ràng việc đến và đi của khối số ống (nguyên thủy duy trì điều phối linh hoạt): đưa ra tay cầm khối số ống + tay cầm từ (fromHandle) + đến tay cầm (toHandle) (khối thẻ thiết bị/khối mã công cộng/khối số ống), đọc các thuộc tính của nó để lấy giá trị từ/đích (có thể được ghi đè bởi từ/đến), ghi thuộc tính đến/từ + vị trí tương đối hai điểm + di chuyển lớp cố định, nhất quán với lệnh cọ vẽ. Hỗ trợ một mục (pipeArrowHandle/fromHandle/toHandle) hoặc hàng loạt (items[]). | pipeArrowHandle, fromHandle, toHandle, from, to, items, setFreezeLayer |
| api_gslc_sync_equipment_data |  | Đồng bộ hóa dữ liệu thuộc tính thiết bị từ máy khách sang bản vẽ CAD | jsonFilePath, syncData, entityHandles |
| api_gslc_sync_instrument_data |  | Đồng bộ hóa dữ liệu thuộc tính thiết bị từ máy khách sang bản vẽ CAD | jsonFilePath, syncData, entityHandles |
| api_gslc_sync_pipe_data |  | Đồng bộ hóa dữ liệu thuộc tính đường ống từ khách hàng sang bản vẽ CAD | jsonFilePath, syncData, entityHandles |
| api_gslc_tag_text_to_block |  | Tự động xác định văn bản thẻ thiết bị thông qua chế độ khoảng trắng (văn bản tên phân tách văn bản thẻ) và chuyển đổi nó thành khối thẻ thiết bị tiêu chuẩn ($ ClaudeCAD-GsLcEquipTag) | searchTolerance, deleteOriginalText, previewOnly, scale |
| api_gslc_text_to_pipe_arrow |  | Chuyển đổi văn bản số ống thành khối số ống, tự động khớp với đường ống liền mạch liên tục gần nhất để xác định vị trí chèn (tìm kiếm hai chiều, ưu tiên phía thông thường, loại trừ đường chấm/đường tín hiệu lồng vào nhau) | layerName, pipeNumPattern, matchTolerance, targetBlockName, deleteOriginalText, previewOnly, scale |
| api_gslc_validate_entity_data |  | Xác minh tính toàn vẹn dữ liệu của thực thể (kiểm tra các thuộc tính bắt buộc, vị trí tương đối, v.v.). Bạn có thể chỉ định phạm vi khung vẽ hoặc xác minh toàn bộ bản vẽ | frameHandle, entityTypes, checkRelativePosition |

## 9. Quản lý Giao dịch (Transaction) (3 lệnh)

| Tên Lệnh API | Tên CAD / NetLoad | Mô Tả Chức Năng | Tham Số Chính |
| :--- | :--- | :--- | :--- |
| api_transaction_begin |  | Khởi tạo một giao dịch chỉnh sửa CAD để hỗ trợ Rollback khi lỗi | - |
| api_transaction_commit |  | Lưu cố định các thay đổi trong giao dịch vào bản vẽ | transactionId |
| api_transaction_rollback |  | Hủy bỏ toàn bộ thay đổi kể từ khi bắt đầu giao dịch | transactionId |

## 10. Quản lý Bản vẽ / Tài liệu (6 lệnh)

| Tên Lệnh API | Tên CAD / NetLoad | Mô Tả Chức Năng | Tham Số Chính |
| :--- | :--- | :--- | :--- |
| api_close_document |  | Đóng bản vẽ hiện hành, tùy chọn lưu hoặc hủy thay đổi | save, filePath |
| api_get_document_info |  | Lấy tên file, đường dẫn, đơn vị INSUNITS và phạm vi bản vẽ | - |
| api_get_drawing_extents |  | Lấy tọa độ góc bao quanh toàn bộ các đối tượng trong bản vẽ | updateExtents |
| api_open_document |  | Mở file DWG/DXF vào không gian làm việc AutoCAD | filePath, readOnly |
| api_purge_database |  | Dọn dẹp các layer, block, kiểu chữ rác không sử dụng | purgeLayers, purgeBlocks, purgeLinetypes, purgeTextStyles, purgeDimStyles |
| api_save_document |  | Lưu bản vẽ hiện hành hoặc lưu sang đường dẫn mới (Save As) | filePath, version |

## 11. Chỉnh sửa thực thể (13 lệnh)

| Tên Lệnh API | Tên CAD / NetLoad | Mô Tả Chức Năng | Tham Số Chính |
| :--- | :--- | :--- | :--- |
| api_bind_xdata |  | Gắn dữ liệu mở rộng XData vào thực thể theo AppName | entityHandle, entityHandles, appName, data |
| api_copy_entity |  | Sao chép đối tượng từ điểm gốc tới điểm đích | entityHandle, entityHandles, fromPoint, toPoint |
| api_delete_entity |  | Xóa bỏ đối tượng khỏi bản vẽ CAD | entityHandle, entityHandles |
| api_mirror_entity |  | Lấy đối xứng đối tượng qua trục 2 điểm | entityHandle, entityHandles, mirrorPoint1, mirrorPoint2, deleteSource |
| api_move_entity |  | Di chuyển đối tượng từ điểm gốc tới điểm đích | entityHandle, entityHandles, fromPoint, toPoint |
| api_offset_entity |  | Tạo đối tượng song song cách một khoảng xác định | entityHandle, distance, sidePoint, side, bothSides |
| api_remove_xdata |  | Xóa dữ liệu XData của AppName trên thực thể | entityHandle, entityHandles, appName, allApps |
| api_rotate_entity |  | Xoay đối tượng quanh điểm chuẩn theo góc độ chỉ định | entityHandle, entityHandles, basePoint, angle |
| api_scale_entity |  | Co giãn tỷ lệ đối tượng quanh điểm chuẩn | entityHandle, entityHandles, basePoint, scaleFactor |
| api_set_entity_color |  | Đổi màu sắc hiển thị của đối tượng theo mã màu ACI | entityHandle, entityHandles, colorIndex, rgb |
| api_set_entity_layer |  | Chuyển đối tượng sang Layer khác | entityHandle, entityHandles, layerName, createIfNotExist |
| api_set_entity_properties |  | Thiết lập hàng loạt thuộc tính của đối tượng (Layer, màu, nét, độ rộng) | entityHandle, entityHandles, properties |
| api_trim_extend_entity |  | Cắt tỉa hoặc kéo dài đối tượng tới biên giao cắt | entityHandle, boundaryHandles, mode, extendMode, pickPoint |

## 12. Layer nâng cao (7 lệnh)

| Tên Lệnh API | Tên CAD / NetLoad | Mô Tả Chức Năng | Tham Số Chính |
| :--- | :--- | :--- | :--- |
| api_create_layer |  | Tạo một Layer mới với màu sắc, kiểu nét và độ dày nét chỉ định | layerName, color, lineType, lineWeight, isOff, isFrozen, isLocked, description |
| api_delete_layer |  | Xóa bỏ Layer rỗng khỏi bản vẽ CAD | layerName, force |
| api_get_layer_info |  | Lấy chi tiết trạng thái, màu sắc, kiểu đường và khóa của Layer | layerName |
| api_list_layers |  | Lấy danh sách tất cả các Layer trong bản vẽ kèm trạng thái | includeOff, includeFrozen, filter |
| api_rename_layer |  | Đổi tên một Layer trong bản vẽ | oldName, newName |
| api_set_current_layer |  | Kích hoạt một Layer làm Layer vẽ mặc định hiện hành | layerName |
| api_set_layer_properties |  | Cập nhật màu sắc, đóng băng (freeze), khóa (lock) cho Layer | layerName, color, lineType, lineWeight, isOff, isFrozen, isLocked, isPlottable, description |

## 13. Điều khiển Viewport (5 lệnh)

| Tên Lệnh API | Tên CAD / NetLoad | Mô Tả Chức Năng | Tham Số Chính |
| :--- | :--- | :--- | :--- |
| api_get_current_view |  | Lấy thông tin tọa độ tâm và kích thước khung nhìn hiện hành | - |
| api_zoom_center |  | Định vị tâm màn hình tại tọa độ chỉ định với độ phóng đại | center, scale, width, height |
| api_zoom_extents |  | Phóng to bao quát toàn bộ nội dung bản vẽ trên màn hình | scale |
| api_zoom_object |  | Phóng to trực tiếp vừa vặn vào đối tượng chỉ định | entityHandle, entityHandles, scale |
| api_zoom_window |  | Phóng to vùng màn hình theo 2 điểm góc chữ nhật | minPoint, maxPoint, margin |

---
**Tổng cộng:** 135 lệnh điều khiển CAD tự động đã được Việt hóa.