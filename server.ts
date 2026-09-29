import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type, ThinkingLevel } from "@google/genai";

async function generateContentWithRetryAndFallback(
  ai: GoogleGenAI,
  model: string,
  parts: any[],
  config: any,
  maxRetries = 3
): Promise<any> {
  let attempt = 0;
  let delay = 1000;

  while (attempt < maxRetries) {
    try {
      console.log(`[Attempt ${attempt + 1}/${maxRetries}] Generating content using ${model}...`);
      
      const currentConfig = { ...config };
      if (!model.startsWith("gemini-3.5")) {
        delete currentConfig.thinkingConfig;
      }
      if (attempt > 0 && currentConfig.thinkingConfig) {
        console.warn(`[Attempt ${attempt + 1}] Disabling thinkingConfig for retry to bypass high-priority queue resource usage.`);
        delete currentConfig.thinkingConfig;
      }

      const response = await ai.models.generateContent({
        model,
        contents: { parts },
        config: currentConfig
      });
      return response;
    } catch (error: any) {
      attempt++;
      
      const errStr = (error?.message || String(error)).toLowerCase();
      const isAuthError = 
        error?.status === 401 || 
        error?.status === 403 || 
        errStr.includes("api key") || 
        errStr.includes("invalid") || 
        errStr.includes("unauthorized") || 
        errStr.includes("permissions");
      const isTransient = !isAuthError;

      if (isTransient && attempt < maxRetries) {
        console.warn(`[Transient Error] Attempt ${attempt} failed: ${error.message || error}. Retrying in ${delay}ms...`);
        await new Promise(resolve => setTimeout(resolve, delay));
        delay *= 2;
      } else {
        throw error;
      }
    }
  }
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Middleware to parse large JSON bodies
  app.use(express.json({ limit: "50mb" }));

  // API route for grading
  app.post("/api/grade", async (req, res) => {
    try {
      const data = req.body;
      const apiKey = (process.env.GEMINI_API_KEY || process.env.API_KEY)?.trim();
      if (!apiKey || apiKey === 'MY_GEMINI_API_KEY') {
        return res.status(401).json({ error: "Chưa cấu hình Gemini API Key. Vui lòng bấm vào biểu tượng bánh răng (Settings) ở góc phải màn hình của AI Studio để nhập API Key của bạn." });
      }
      const ai = new GoogleGenAI({
        apiKey,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build',
          }
        }
      });
      const model = "gemini-3.5-flash";

      const SYSTEM_INSTRUCTION = `
ROLE: Chuyên gia khảo thí cao cấp, cộng sự đắc lực của Thầy Vinh.
NHIỆM VỤ: Phân tích dữ liệu Sheet (TSV/CSV), chấm điểm (TỐI ĐA 10) và viết nhận xét cá nhân hóa theo ĐÚNG 4 Ý CHI TIẾT (tuyệt đối không nhận xét liên quan đến camera).

QUY TẮC CHẤM ĐIỂM VÀ XẾP HẠNG:
1. ĐIỂM TỐI ĐA: Tổng điểm tuyệt đối không được vượt quá 10.
2. CỘNG THƯỞNG (Tương tác): Nếu có tên trong [DANH SÁCH KHEN], hãy cộng 1 điểm thưởng vào điểm thô.
3. GIỚI HẠN: Sau khi cộng tất cả các điểm, nếu tổng điểm vượt quá 10, hãy giữ nguyên ở mức 10. Điểm tối thiểu là 0.
4. NHÁP NGẦM: Tái cấu trúc bài làm từ ảnh/văn bản, đối chiếu bước làm chi tiết với biểu mẫu chấm, xác minh lỗi sai ít nhất 2 lần để đảm bảo chính xác 100%.
5. XẾP HẠNG (rank): Dựa vào điểm số cuối cùng (sau khi cộng điểm):
   - Điểm < 5: Không đạt
   - Điểm <= 7: Khuyến khích
   - Điểm = 8: Ba
   - Điểm = 9: Nhì
   - Điểm = 10: Nhất

QUY TẮC NHẬN DIỆN DỮ LIỆU:
- TSV/TAB: Nhận diện cột qua Tab (\\t). Thứ tự mặc định: [Thời gian nộp, Họ tên, Tên, Bài làm].
- Xuống dòng: Gom toàn bộ nội dung bài làm của một học sinh vào một khối duy nhất cho học sinh đó.

QUY TẮC NHẬN XÉT (BẮT BUỘC GIỮ ĐÚNG THỨ TỰ 4 Ý TRONG MẢNG FEEDBACK - TUYỆT ĐỐI KHÔNG ICON/EMOJI - KHÔNG NHẬN XÉT LIÊN QUAN ĐẾN CAMERA):
Lưu ý quan trọng: Để tránh việc nhận xét bị lặp đi lặp lại giống như máy móc, bạn PHẢI DIỄN ĐẠT LINH HOẠT, SỬ DỤNG TỪ ĐỒNG NGHĨA, CẤU TRÚC CÂU ĐA DẠNG cho các Ý 2, Ý 3. Tuy nhiên, VẪN PHẢI GIỮ NGUYÊN LOGIC CỐT LÕI VÀ THÔNG ĐIỆP CHÍNH.

Ý 1 (Nội dung bài làm - Kết quả):
- Nếu ĐÚNG HẾT: Con làm bài đạt kết quả rất tốt, rất đáng được khen ngợi.
- Nếu LÀM SAI: [Thông báo cụ thể câu sai & đáp án đúng kèm lý do - chỉ rõ lỗi sai ở dòng/bước nào].

Ý 2 (Nội dung bài làm - Hành động) - [DIỄN ĐẠT LINH HOẠT TỪ NGỮ]:
- Nếu ĐÚNG HẾT (Logic cốt lõi): Khuyến khích ba mẹ tiếp tục động viên, đồng hành để con thêm tự tin, mạnh dạn hơn trong học tập.
- Nếu LÀM SAI (Logic cốt lõi): Nhờ phụ huynh hỗ trợ, nhắc nhở con xem lại video/slide bài giảng và làm lại bài thi đầu ra cho đến khi đúng với đáp án thầy đã chữa.

Ý 3 (Tương tác) - [DIỄN ĐẠT LINH HOẠT TỪ NGỮ]:
- CÓ TÊN trong DANH SÁCH KHEN (Logic cốt lõi): Tuyên dương tinh thần học tập sôi nổi, con hăng hái phát biểu, tích cực nhắn tin tương tác với thầy, thông báo thầy cộng con thêm 1 điểm.
- KHÔNG CÓ TÊN (Logic cốt lõi): Nhắc nhở nhẹ nhàng việc hôm nay con còn trầm, ít tham gia nhắn tin trả lời bài, khích lệ con mạnh dạn và tương tác nhiều hơn vào buổi sau.

Ý 4 (Lời chúc & Lời ngỏ):
- Cố gắng phát huy kết quả này, thầy tin con sẽ ngày càng tiến bộ. (Thay đổi linh hoạt, chân thành để phụ huynh thấy sự quan tâm từ Thầy Vinh).

YÊU CẦU ĐẶC BIỆT:
- TUYỆT ĐỐI KHÔNG DÙNG ICON/EMOJI TRONG NHẬN XÉT.
- TUYỆT ĐỐI KHÔNG ĐƯỢC ĐỀ CẬP HOẶC NHẬN XÉT BẤT KỲ ĐIỀU GÌ LIÊN QUAN ĐẾN CAMERA/GÓC QUAY/CAM BẬT/CAM TẮT.
- Trả KẾT QUẢ THEO ĐÚNG THỨ TỰ HỌC SINH CÓ TRONG DỮ LIỆU ĐẦU VÀO, tuyệt đối không được tự ý đảo lộn hay sắp xếp lại tên học sinh.
- Báo cáo lỗi dữ liệu vào "validationWarnings" nếu dòng dữ liệu bị thiếu thông tin hoặc sai cấu trúc.
`;

      const parts: any[] = [
        { text: `BIỂU MẪU CHẤM ĐIỂM CHI TIẾT CỦA THẦY VINH: ${data.markingGuide}` },
        { text: `--- DỮ LIỆU ĐẦU VÀO ---` },
        { text: `Dữ liệu Sheet (TSV/CSV): ${data.sheetData}` },
        { text: `Danh sách Bật Cam (Rõ - Thấy màn hình/vở): ${data.camVisibleList}` },
        { text: `Danh sách Không Bật Cam/Mờ: ${data.camHiddenList}` },
        { text: `Danh sách Khen tương tác (+1đ): ${data.praiseList}` }
      ];

      if (data.testImage) {
        parts.unshift({
          inlineData: {
            mimeType: "image/jpeg",
            data: data.testImage.split(',')[1]
          }
        });
      }

      const config = {
        systemInstruction: SYSTEM_INSTRUCTION,
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            answerKey: { type: Type.STRING, description: "Lời giải chi tiết và đáp án chuẩn AI tự giải từ đề bài." },
            validationWarnings: { type: Type.ARRAY, items: { type: Type.STRING } },
            results: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  studentName: { type: Type.STRING, description: "Họ và tên học sinh" },
                  firstName: { type: Type.STRING, description: "Tên học sinh" },
                  studentAnswer: { type: Type.STRING, description: "Bài làm của học sinh" },
                  score: { type: Type.INTEGER, description: "Điểm số sau khi cộng tất cả các khoản thưởng, tối đa 10." },
                  submissionTime: { type: Type.STRING, description: "Thời gian nộp bài" },
                  rank: { type: Type.STRING, description: "Xếp hạng" },
                  feedback: { 
                    type: Type.ARRAY, 
                    items: { type: Type.STRING },
                    description: "Mảng chứa đúng 4 chuỗi nhận xét theo đúng thứ tự template của Thầy Vinh (Ý 1, Ý 2, Ý 3, Ý 4). Ý 4 chứa lời chúc & lời ngỏ."
                  }
                },
                required: ["studentName", "firstName", "studentAnswer", "score", "submissionTime", "rank", "feedback"]
              }
            }
          },
          required: ["answerKey", "results"]
        }
      };

      let response;
      try {
        response = await generateContentWithRetryAndFallback(ai, model, parts, config, 3);
      } catch (firstError: any) {
        console.warn("All retry attempts failed for primary model gemini-3.5-flash. Attempting fallback model gemini-3.1-flash-lite...");
        const fallbackConfig = { ...config };

        try {
          response = await generateContentWithRetryAndFallback(ai, "gemini-3.1-flash-lite", parts, fallbackConfig, 2);
        } catch (fallbackError: any) {
          console.warn("Fallback gemini-3.1-flash-lite also failed. Attempting alternative model gemini-flash-latest...");
          try {
            response = await generateContentWithRetryAndFallback(ai, "gemini-flash-latest", parts, fallbackConfig, 2);
          } catch (lastError: any) {
            console.error("All fallback models failed:", lastError);
            throw firstError; // Throw the original error which has the most accurate details for the primary model
          }
        }
      }

      const report = JSON.parse(response.text || "{}");
      res.json(report);
    } catch (error: any) {
      console.error("Gemini API Error:", error);
      
      let errorMessage = "Lỗi không xác định khi chấm bài.";
      if (error?.status === 401 || error?.status === 400 || error?.message?.includes("API key not valid")) {
        errorMessage = "API Key không hợp lệ. Vui lòng kiểm tra lại cấu hình Gemini API Key trong menu Settings.";
      } else if (error?.status === 429 || error?.message?.includes("RESOURCE_EXHAUSTED") || error?.message?.includes("quota")) {
        errorMessage = "API Key đã hết hạn mức (Quota/Credits). Vui lòng nạp thêm tín dụng hoặc sử dụng API Key khác.";
      } else if (error?.message) {
        errorMessage = error.message;
      }

      res.status(500).json({ error: errorMessage });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*all', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
