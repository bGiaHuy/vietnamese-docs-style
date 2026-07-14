---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: 7cea322cf7c584557d31710ecddbcd34_f5a0bbd7610711f1832e5254006c9bbf
    ReservedCode1: dcvBgJsWPTL4JtTql6OYtfisPALQ5zb9SivCSRxPb79TYXZ+4ettjw6wTscZIAWOXmjI/tM4awrkVXZF1/0BYMbvE5KynNsrgIXg0r/sXbrZna2VENnyJ1h2n3k5FEKuW3o1JSrNthjglGmheHDPdKKoIznAnAlNtWTqtcuK4NhJ6cGZSfO46/yutKE=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: 7cea322cf7c584557d31710ecddbcd34_f5a0bbd7610711f1832e5254006c9bbf
    ReservedCode2: dcvBgJsWPTL4JtTql6OYtfisPALQ5zb9SivCSRxPb79TYXZ+4ettjw6wTscZIAWOXmjI/tM4awrkVXZF1/0BYMbvE5KynNsrgIXg0r/sXbrZna2VENnyJ1h2n3k5FEKuW3o1JSrNthjglGmheHDPdKKoIznAnAlNtWTqtcuK4NhJ6cGZSfO46/yutKE=
---

# Phân tích Ngôn ngữ Hành chính Việt Nam

> **Nguồn dữ liệu**: 6.010 văn bản từ train_data (công văn, báo cáo, biên bản, đề tài, hợp đồng, giáo án)
> **Phương pháp**: Phân tích n-gram, cấu trúc câu, cụm từ cố định, phát hiện lỗi sai
> **Mục đích**: Bổ sung phần ngôn ngữ cho SKILL_v4_NĐ30 — giúp AI sinh văn bản tự nhiên, đúng phong cách hành chính Việt Nam

---

## I. Tổng quan dữ liệu

| Loại văn bản | Số lượng | Tỷ lệ | Đặc điểm |
|-------------|---------|------|----------|
| Quyết định | 1.482 | 24,7% | Văn bản cá biệt, Điều khoản, hiệu lực thi hành |
| Biên bản | 1.000 | 16,6% | Tổng hợp, có Quốc hiệu + Tiêu ngữ + ký tên |
| Giáo án | 1.000 | 16,6% | Cấu trúc 5 hoạt động, mục tiêu – nội dung – sản phẩm |
| Công văn | 501 | 8,3% | Hành chính giao dịch, kính gửi, trích yếu |
| Sắc lệnh | 483 | 8,0% | Văn bản lịch sử, "RA SẮC LỆNH" |
| Thông tư | 456 | 7,6% | Hướng dẫn thi hành, nhiều Điều khoản |
| Chỉ thị | 325 | 5,4% | Mệnh lệnh hành chính, "CHỈ THỊ" |
| Nghị định | 303 | 5,0% | Quy phạm chi tiết, Chương/Mục/Điều |
| Nghị quyết | 248 | 4,1% | Quyết nghị tập thể |
| Khác | 212 | 3,5% | Thông tư liên tịch, Pháp lệnh, Luật, Lệnh |

---

## II. Cấu trúc điển hình theo loại văn bản

### 2.1 Văn bản quy phạm (Nghị định, Thông tư, Nghị quyết, Quyết định, Chỉ thị)

**Cấu trúc lặp lại (>60% văn bản):**

```
[QUỐC HIỆU + TIÊU NGỮ]            ← 11-26% (thường bị cắt trong mẫu)
[TÊN CƠ QUAN BAN HÀNH]            ← 100% luật cũ, NĐ30 yêu cầu
[SỐ/KÝ HIỆU]                       ← ngầm trong metadata

TÊN LOẠI VĂN BẢN                  ← IN HOA, ĐẬM
Trích yếu nội dung                 ← in thường, đậm

[CƠ QUAN BAN HÀNH lặp lại]
Căn cứ [văn bản pháp lý]...;       ← 56-64% có; mỗi dòng 1 căn cứ
Căn cứ [văn bản pháp lý]...;
...
Theo đề nghị của...;              ← có trong ~40% văn bản

QUYẾT ĐỊNH:                        ← 28-39% văn bản

Điều 1. [nội dung]                ← 62-75% có Điều khoản
Điều 2. [nội dung]
...

Điều X. [hiệu lực/trách nhiệm]   ← thường là Điều cuối
  Quyết định này có hiệu lực...
  Các ông/bà... có trách nhiệm thi hành...

[NƠI NHẬN]                        ← 0% trong mẫu (bị cắt)
[KÝ TÊN/ĐÓNG DẤU]                 ← 1-4% (bị cắt)
```

### 2.2 Biên bản (1.000 mẫu)

**Cấu trúc gần như cố định (100%):**

```
CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM          ← 98%
Độc lập - Tự do - Hạnh phúc                  ← 97.5%
--------------------------------------------------  ← 97.5%

BIÊN BẢN [LOẠI]                                ← IN HOA, ĐẬM

Căn cứ [văn bản pháp lý];                      ← 99.5%
Căn cứ [Điều lệ/Quy chế];

Hôm nay, vào lúc [giờ], ngày [ngày tháng năm], ← 100%
tại [địa điểm], [Đơn vị] tiến hành họp
với thành phần như sau:

I. THÀNH PHẦN THAM DỰ                           ← 100%
  1. [Họ tên] - [Chức vụ]
  ...

II. NỘI DUNG                                     ← 100%
1. [Mục 1]
   - Tình hình thực tế: Đã triển khai đúng kế hoạch...
   - Ý kiến thảo luận: Các thành viên tham dự đều nhất trí...
   - Kết luận: [Phê duyệt / Ghi nhận / Đề nghị / Tiếp tục theo dõi]

III. KẾT LUẬN CHUNG                              ← 100%
  - Cuộc họp kết thúc vào lúc [giờ] cùng ngày.
  - Các bên thống nhất thực hiện theo đúng các nội dung đã thông qua.
  - Biên bản này được lập thành [N] bản, có giá trị pháp lý như nhau.

                              THƯ KÝ            CHỦ TRÌ         ← 97.5%
                                (Ký, ghi rõ họ tên)   (Ký, ghi rõ họ tên)
```

**20 loại biên bản phổ biến trong dữ liệu:**
Biên bản họp chi bộ, họp hội đồng quản trị, hòa giải tranh chấp, sinh hoạt lớp, làm việc với đối tác, họp hội đồng thi đua, họp ban chấp hành công đoàn, thanh lý hợp đồng, đối thoại, kiểm tra, bàn giao, nghiệm thu, xử lý vi phạm, họp phụ huynh, họp giao ban, họp hội đồng kỷ luật, họp tổng kết, họp hội đồng nhân dân, tiếp công dân, họp chi ủy.

### 2.3 Giáo án (1.000 mẫu)

**Cấu trúc cố định (100%):**

```
GIÁO ÁN [MÔN HỌC] - LỚP [X]
Bài: [Tên bài]
Thời lượng: [N] tiết
Ngày soạn: [YYYY-MM-DD]

I. MỤC TIÊU
1. Kiến thức:
  - Học sinh [động từ năng lực] được [nội dung].
2. Năng lực:
  - Năng lực [giao tiếp và hợp tác / tìm hiểu tự nhiên / giải quyết vấn đề / tự chủ]
3. Phẩm chất:
  - [Chăm chỉ / Nhân ái / Yêu nước / Trung thực / Trách nhiệm]

II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
1. Giáo viên: SGK, SGV, giáo án, máy chiếu, bảng phụ, phiếu học tập.
2. Học sinh: SGK, vở ghi, đồ dùng học tập.

III. TIẾN TRÌNH DẠY HỌC
Hoạt động 1: Khởi động ([N] phút)
Hoạt động 2: Hình thành kiến thức ([N] phút)
Hoạt động 3: Luyện tập ([N] phút)
Hoạt động 4: Vận dụng ([N] phút)
Hoạt động 5: Tổng kết ([N] phút)

Mỗi hoạt động có cấu trúc con:
  a) Mục tiêu: ...
  b) Nội dung: ...
  c) Sản phẩm: Câu trả lời của HS, [phiếu học tập/bài tập nhóm/sơ đồ tư duy].
  d) Tổ chức thực hiện:
    Bước 1 - Chuyển giao nhiệm vụ: GV nêu yêu cầu, giao nhiệm vụ cho HS.
    Bước 2 - Thực hiện nhiệm vụ: HS [quan sát/đọc SGK/làm việc nhóm]. GV quan sát, hỗ trợ.
    Bước 3 - Báo cáo, thảo luận: [N] HS/nhóm trình bày. Các HS/nhóm khác nhận xét, bổ sung.
    Bước 4 - Kết luận, nhận định: GV nhận xét, đánh giá và chốt kiến thức trọng tâm.

RÚT KINH NGHIỆM SAU TIẾT DẠY:
............................................................................
```

---

## III. Cách dùng từ đặc trưng

### 3.1 Động từ hành động phổ biến (theo tần suất)

| STT | Động từ | Tần suất | Ngữ cảnh điển hình |
|-----|---------|---------|-------------------|
| 1 | **tổ chức** | Rất cao | "tổ chức thực hiện", "tổ chức kiểm tra" |
| 2 | **thực hiện** | Rất cao | "có trách nhiệm thực hiện", "triển khai thực hiện" |
| 3 | **quản lý** | Cao | "quản lý và sử dụng", "quản lý nhà nước về..." |
| 4 | **hướng dẫn** | Cao | "hướng dẫn thi hành", "hướng dẫn thực hiện" |
| 5 | **ban hành** | Cao | "ban hành kèm theo", "ban hành quy định" |
| 6 | **xây dựng** | Cao | "xây dựng kế hoạch", "xây dựng và phát triển" |
| 7 | **kiểm tra** | Trung bình | "kiểm tra, giám sát", "thanh tra, kiểm tra" |
| 8 | **phát triển** | Trung bình | "phát triển kinh tế - xã hội" |
| 9 | **phối hợp** | Trung bình | "phối hợp với các cơ quan" |
| 10 | **đánh giá** | Trung bình | "đánh giá kết quả", "nhận xét, đánh giá" |

### 3.2 Động từ tình thái (modal verbs)

| Động từ | Mức độ | Ví dụ |
|---------|--------|-------|
| **phải** | Bắt buộc | "Các đơn vị phải thực hiện nghiêm túc" |
| **cần** | Khuyến nghị mạnh | "Cần tăng cường công tác kiểm tra" |
| **yêu cầu** | Chỉ đạo | "Yêu cầu các bộ, ngành..." |
| **đề nghị** | Kiến nghị | "Đề nghị ... xem xét, giải quyết" |
| **không được** | Cấm | "Không được tự ý thay đổi" |
| **có thể** | Cho phép | "Có thể áp dụng linh hoạt" |
| **nên** | Khuyến nghị nhẹ | "Nên tham khảo ý kiến chuyên gia" |
| **cấm** | Tuyệt đối | "Nghiêm cấm mọi hành vi..." |

### 3.3 Từ nối chuyển tiếp (connectors)

| Từ nối | Tần suất | Chức năng |
|--------|---------|-----------|
| **đồng thời** | Rất cao | Bổ sung ý song song |
| **trước hết / thứ nhất** | Trung bình | Mở đầu liệt kê |
| **thứ hai / thứ ba** | Trung bình | Liệt kê tuần tự |
| **cuối cùng / sau cùng** | Trung bình | Kết thúc liệt kê |
| **ngoài ra** | Trung bình | Bổ sung ý phụ |
| **tuy nhiên** | Trung bình | Đối lập |
| **mặt khác** | Trung bình | Góc nhìn khác |
| **vì vậy / do đó** | Thấp | Kết luận logic |
| **bên cạnh đó** | Thấp | Bổ sung |

### 3.4 Cụm từ cố định (collocations)

| Loại văn bản | Cụm từ cố định | Ghi chú |
|-------------|---------------|---------|
| Văn bản luật | *Căn cứ [tên luật] số... ngày...;* | Mỗi căn cứ 1 dòng, cuối `;`, dòng cuối `.` |
| Văn bản luật | *Theo đề nghị của [cơ quan]...;* | Thường ở dòng cuối trước QUYẾT ĐỊNH |
| Văn bản luật | *QUYẾT ĐỊNH:* | Từ khóa chuyển từ căn cứ sang nội dung |
| Văn bản luật | *Quyết định này có hiệu lực [kể từ/từ ngày]...* | Điều khoản thi hành |
| Văn bản luật | *...có trách nhiệm thi hành Quyết định này.* | Điều khoản trách nhiệm |
| Biên bản | *Căn cứ [văn bản]; Căn cứ [quy chế];* | Mở đầu biên bản |
| Biên bản | *Hôm nay, vào lúc... ngày... tại...* | Thời gian, địa điểm |
| Biên bản | *Cuộc họp kết thúc vào lúc... cùng ngày.* | Kết luận chung |
| Biên bản | *Các bên thống nhất thực hiện theo đúng các nội dung đã thông qua.* | Kết luận chung |
| Biên bản | *Biên bản này được lập thành [N] bản, có giá trị pháp lý như nhau.* | Kết thúc |
| Giáo án | *Học sinh [vận dụng/trình bày/phân tích] được...* | Mục tiêu kiến thức |
| Giáo án | *GV nêu yêu cầu, giao nhiệm vụ cho HS.* | Bước 1 mọi hoạt động |
| Giáo án | *GV quan sát, hỗ trợ kịp thời.* | Bước 2 mọi hoạt động |
| Giáo án | *Các HS/nhóm khác nhận xét, bổ sung.* | Bước 3 mọi hoạt động |
| Giáo án | *GV nhận xét, đánh giá và chốt kiến thức trọng tâm.* | Bước 4 mọi hoạt động |

---

## IV. Cấu trúc câu điển hình

### 4.1 Văn bản quy phạm pháp luật

**A. Câu mở đầu phần căn cứ:**
```
Căn cứ [Tên luật/văn bản] số [XX]/[năm]/[cơ quan] ngày [dd] tháng [mm] năm [yyyy];
```

**B. Câu chuyển tiếp hành động:**
```
Sau khi xem xét đề nghị của...;           ← dùng cho Quyết định
Theo đề nghị của...;                      ← dùng cho Nghị quyết
Xét đề nghị của...;                       ← dùng cho Chỉ thị
```

**C. Câu khởi đầu phần quyết định (nội dung):**
```
Điều 1. [Phạm vi điều chỉnh / Đối tượng áp dụng]
Điều 1. Ban hành kèm theo [Quyết định/Nghị định] này...
Điều 1. Nay [quy định/thành lập/cho phép/bãi bỏ]...
```

**D. Câu về hiệu lực (luôn là một trong các Điều cuối):**
```
[Văn bản] này có hiệu lực thi hành kể từ ngày ký.
[Văn bản] này có hiệu lực kể từ ngày [dd] tháng [mm] năm [yyyy].
[Văn bản] này có hiệu lực sau [N] ngày kể từ ngày ký.
```

**E. Câu về trách nhiệm thi hành:**
```
Các Bộ trưởng, Thủ trưởng cơ quan ngang Bộ... có trách nhiệm thi hành [văn bản] này.
Bộ trưởng Bộ [Tên bộ] chịu trách nhiệm thi hành [văn bản] này.
Chủ tịch UBND các tỉnh, thành phố... chịu trách nhiệm tổ chức thực hiện...
```

**F. Câu trong văn bản sắc lệnh (lịch sử):**
```
Chiểu theo Sắc lệnh ngày [dd-mm-yyyy]...
Sau khi xét đề nghị của...;
RA SẮC LỆNH:
Điều thứ nhất: Kể từ ngày ký sắc lệnh này...
Các ông Bộ trưởng... sẽ tuỳ chức vụ mà thi hành Sắc lệnh này.
```

### 4.2 Biên bản

**A. Câu mở đầu nội dung:**
```
1. [Tên mục]
   - Tình hình thực tế: Đã triển khai đúng kế hoạch đề ra, đạt [N]% chỉ tiêu.
   - Ý kiến thảo luận: Các thành viên tham dự đều nhất trí với nội dung báo cáo.
   - Kết luận: [Phê duyệt / Ghi nhận / Đề nghị bổ sung / Tiếp tục theo dõi] nội dung này.
```

**B. 4 loại kết luận phổ biến:**
| Kết luận | Tần suất | Ngữ cảnh |
|----------|---------|----------|
| Phê duyệt nội dung này | Cao nhất | Nội dung được thông qua |
| Ghi nhận nội dung này | Trung bình | Nội dung cần ghi nhận, không cần phê duyệt |
| Đề nghị bổ sung thêm nội dung này | Thấp | Cần thêm thông tin |
| Tiếp tục theo dõi nội dung này | Thấp | Chưa có kết luận cuối cùng |

### 4.3 Giáo án

**A. Động từ mục tiêu kiến thức (chuẩn):**
```
Học sinh vận dụng được [khái niệm] về [chủ đề].
Học sinh trình bày được [ý nghĩa/vai trò] của [chủ đề].
Học sinh phân tích được [nội dung/đặc điểm] của [chủ đề].
Học sinh nắm bắt được kiến thức về [chủ đề].
Học sinh mở rộng kiến thức về [chủ đề].
```

**B. 3 mẫu tổ chức lớp học:**
| Mẫu | Nội dung Bước 1 |
|-----|-----------------|
| Cá nhân | *GV tổ chức cho HS làm việc cá nhân để tìm hiểu nội dung bài học.* |
| Nhóm | *GV chia lớp thành các nhóm để tìm hiểu nội dung bài học.* |
| Gợi mở | *GV đặt câu hỏi gợi mở để tìm hiểu nội dung bài học.* |

**C. 4 loại sản phẩm đầu ra:**
| Sản phẩm | Mẫu câu |
|----------|---------|
| Câu trả lời | *Câu trả lời của HS, phiếu học tập đã hoàn thành.* |
| Bài tập nhóm | *Câu trả lời của HS, bài tập nhóm.* |
| Sơ đồ tư duy | *Câu trả lời của HS, sơ đồ tư duy.* |
| Phần trình bày | *Câu trả lời của HS, phần trình bày của HS.* |

---

## V. Lỗi thường gặp và cần tránh

### 5.1 Lỗi về cấu trúc

| STT | Lỗi | Tỷ lệ mắc | Cách khắc phục |
|-----|-----|-----------|---------------|
| 1 | **Thiếu "Nơi nhận"** | ~100% mẫu (bị cắt) | Luôn thêm phần Nơi nhận cuối văn bản |
| 2 | **Thiếu phần "Căn cứ"** | 36-57% | Phải có ít nhất 1 căn cứ pháp lý |
| 3 | **Thiếu Quốc hiệu** | 85-91% văn bản luật | Góc trên bên PHẢI |
| 4 | **Dấu cuối căn cứ sai** | ~10% | Dòng cuối phải là dấu chấm (.) không phải (;) |
| 5 | **Tên loại văn bản không IN HOA** | 44-54% | QUYẾT ĐỊNH / NGHỊ QUYẾT phải in hoa |

### 5.2 Lỗi về ngôn ngữ

| STT | Lỗi | Mô tả | Cách khắc phục |
|-----|-----|-------|---------------|
| 6 | **Thiếu động từ hành chính** | Dùng từ ngữ đời thường thay vì từ hành chính | Dùng "tổ chức thực hiện" thay vì "làm", "kiểm tra, giám sát" thay vì "coi" |
| 7 | **Dùng từ nối không đúng** | "Và" lặp quá nhiều thay vì "đồng thời", "bên cạnh đó" | Dùng đa dạng từ nối phù hợp |
| 8 | **Sai mẫu câu cố định** | Không dùng đúng cụm "Quyết định này có hiệu lực..." | Học thuộc các cụm mẫu ở mục III.4 |
| 9 | **Giáo án thiếu mẫu 4 bước** | Không có đủ Bước 1-2-3-4 trong mỗi hoạt động | Mỗi hoạt động phải có đủ 4 bước chuẩn |
| 10 | **Biên bản thiếu kết luận mẫu** | Không có "Các bên thống nhất thực hiện..." | Thêm đầy đủ 3 dòng kết luận chung |

### 5.3 Lỗi về cách hành văn

| STT | Lỗi | Ví dụ sai | Ví dụ đúng |
|-----|-----|----------|-----------|
| 11 | Câu quá dài, khó đọc | Một câu dài >50 từ không ngắt | Tách thành 2-3 câu ngắn hơn |
| 12 | Thiếu tính trang trọng | "Các bạn phải làm theo" | "Các đơn vị có trách nhiệm thực hiện" |
| 13 | Lạm dụng bị động | "Được thực hiện bởi..." | "Thực hiện bởi..." hoặc chủ động |
| 14 | Không nhất quán cách xưng hô | Lúc dùng "cơ quan", lúc dùng "đơn vị" | Chọn 1 cách và dùng xuyên suốt |

---

## VI. Bảng tham chiếu nhanh: Công thức câu theo loại văn bản

### 6.1 Quyết định / Nghị quyết / Chỉ thị

| Vị trí | Công thức | Ví dụ |
|--------|----------|-------|
| Căn cứ 1 | `Căn cứ [Luật] số [XX]/[năm]/[QH] ngày [dd]/[mm]/[yyyy];` | *Căn cứ Luật Tổ chức Chính phủ số 76/2015/QH13 ngày 19/6/2015;* |
| Căn cứ cuối | `Theo đề nghị của [cơ quan]...;` | *Theo đề nghị của Bộ trưởng Bộ Tài chính;* |
| Chuyển tiếp | `QUYẾT ĐỊNH:` | *QUYẾT ĐỊNH:* |
| Điều 1 | `Điều 1. [Động từ] [đối tượng] [nội dung].` | *Điều 1. Ban hành kèm theo Quyết định này...* |
| Hiệu lực | `[Văn bản] này có hiệu lực [thời điểm].` | *Quyết định này có hiệu lực thi hành kể từ ngày ký.* |
| Trách nhiệm | `[Chủ thể] có trách nhiệm thi hành [văn bản] này.` | *Các Bộ trưởng, Thủ trưởng cơ quan...có trách nhiệm thi hành...* |

### 6.2 Biên bản

| Vị trí | Công thức |
|--------|----------|
| Mở đầu | `Căn cứ [văn bản];` × 2 |
| Thời gian | `Hôm nay, vào lúc [giờ], ngày [dd]/[mm]/[yyyy], tại [địa điểm]...` |
| Thành phần | `1. [Họ tên] - [Chức vụ]` |
| Kết luận chung | `Cuộc họp kết thúc vào lúc [giờ] cùng ngày.` |
| | `Các bên thống nhất thực hiện theo đúng các nội dung đã thông qua.` |
| | `Biên bản này được lập thành [N] bản, có giá trị pháp lý như nhau.` |
| Ký tên | `THƯ KÝ` (trái) — `CHỦ TRÌ` (phải) |

### 6.3 Giáo án

| Vị trí | Công thức |
|--------|----------|
| Mục tiêu | `Học sinh [động từ] được [nội dung] về [chủ đề].` |
| Hoạt động | `Hoạt động [N]: [Tên] ([số phút] phút)` |
| Bước 1 | `GV nêu yêu cầu, giao nhiệm vụ cho HS.` |
| Bước 2 | `HS [hành động]. GV quan sát, hỗ trợ kịp thời.` |
| Bước 3 | `[N] HS/nhóm trình bày kết quả. Các HS/nhóm khác nhận xét, bổ sung.` |
| Bước 4 | `GV nhận xét, đánh giá và chốt kiến thức trọng tâm.` |

---

## VII. Đặc điểm nhận diện nhanh từng loại văn bản

| Loại | Dấu hiệu nhận diện | Từ khóa đặc trưng |
|------|-------------------|------------------|
| **Nghị quyết** | "QUYẾT NGHỊ:" thay vì "QUYẾT ĐỊNH:" | *Quốc hội, Hội đồng nhân dân, Hội đồng Bộ trưởng* |
| **Quyết định** | "QUYẾT ĐỊNH:" + Điều 1, 2, 3... | *có trách nhiệm thi hành, có hiệu lực kể từ* |
| **Chỉ thị** | "CHỈ THỊ" + danh sách đánh số 1, 2, 3... | *chỉ thị, yêu cầu, các đơn vị phải* |
| **Thông tư** | "QUYẾT ĐỊNH:" + nhiều Điều + Phụ lục | *theo quy định, hướng dẫn, phụ lục kèm theo* |
| **Nghị định** | "NGHỊ ĐỊNH:" + Chương/Mục/Điều | *Chính phủ, Hội đồng Bộ trưởng, quy định chi tiết* |
| **Sắc lệnh** | "RA SẮC LỆNH:" + "Điều thứ nhất:" | *Chủ tịch nước, Chính phủ lâm thời, chiểu theo* |
| **Công văn** | "Kính gửi:" + nội dung ngắn | *kính gửi, trân trọng, đề nghị phối hợp* |
| **Biên bản** | "BIÊN BẢN [LOẠI]" + I/II/III | *thành phần tham dự, chủ trì, thư ký, kết luận* |
| **Giáo án** | "GIÁO ÁN [MÔN]" + I/II/III + Hoạt động | *mục tiêu, năng lực, phẩm chất, tiến trình dạy học* |

---

> **Sử dụng**: Tài liệu này bổ sung cho SKILL_v4_NĐ30.md. Khi sinh văn bản, kết hợp cả hai: Mục I–VII của SKILL_v4_NĐ30 cho **thể thức** (font, cỡ, lề, vị trí); Mục II–VI của tài liệu này cho **ngôn ngữ** (cách dùng từ, cấu trúc câu, cụm từ cố định). Luôn dùng checklist Mục VI của SKILL_v4_NĐ30.md trước khi xuất file.
*（内容由AI生成，仅供参考）*
