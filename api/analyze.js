const SYSTEM_PROMPT = `너는 대한민국 학교회계 지출품의 업무를 돕는 견적서 분석 전문가다.
업로드된 견적서를 문맥적으로 읽고 실제 구매 품목만 추출한다.

반드시 지킬 규칙:
1. 합계, 총계, 소계, 공급가액 합계, 부가세 합계, 세액, 배송비 안내, 업체정보, 계좌번호, 대표자, 주소, 비고, 견적 유효기간은 품목으로 만들지 않는다.
2. 각 실제 품목마다 내용, 규격, 수량, 부가세 포함 예상단가를 추출한다.
3. 견적서 단가가 부가세 별도이면 부가세를 포함한 단가로 변환한다.
4. 품목별 세액이 따로 있으면 그 품목의 공급가액과 세액을 합쳐 수량으로 나눈 값을 예상단가로 한다.
5. 문서 하단에 부가세만 일괄 표시되면 품목별 공급가액 비율대로 합리적으로 배분한다.
6. 이미 VAT 포함, 부가세 포함이면 다시 더하지 않는다.
7. 면세 품목은 부가세를 추가하지 않는다.
8. 예상금액은 수량 × 예상단가이며, 가능한 경우 견적서 최종 결제 합계와 일치하도록 검산한다.
9. 확신이 낮은 값은 임의로 꾸미지 말고 warnings에 설명한다.
10. 한국어 품목명과 규격을 원문에 가깝게 유지한다.`;

const schema = {
  type: 'object',
  additionalProperties: false,
  properties: {
    supplier: { type: 'string' },
    quoteTotal: { type: 'number' },
    items: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        properties: {
          name: { type: 'string' },
          spec: { type: 'string' },
          quantity: { type: 'number' },
          unitPrice: { type: 'number' },
          amount: { type: 'number' },
          confidence: { type: 'number' }
        },
        required: ['name', 'spec', 'quantity', 'unitPrice', 'amount', 'confidence']
      }
    },
    warnings: { type: 'array', items: { type: 'string' } }
  },
  required: ['supplier', 'quoteTotal', 'items', 'warnings']
};

function extractOutputText(data) {
  if (typeof data.output_text === 'string') return data.output_text;
  for (const item of data.output || []) {
    for (const content of item.content || []) {
      if (content.type === 'output_text' && typeof content.text === 'string') return content.text;
    }
  }
  return '';
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'POST 요청만 지원합니다.' });
  if (!process.env.OPENAI_API_KEY) return res.status(500).json({ error: 'Vercel 환경변수 OPENAI_API_KEY가 설정되지 않았습니다.' });

  try {
    const { text = '', images = [], fileName = '' } = req.body || {};
    if (!text && (!Array.isArray(images) || !images.length)) {
      return res.status(400).json({ error: '분석할 견적서 내용이 없습니다.' });
    }

    const userContent = [{
      type: 'input_text',
      text: `파일명: ${fileName}\n\n다음 견적서에서 실제 품목만 추출하세요.\n${text || '(문자 추출 결과 없음. 첨부 이미지를 분석하세요.)'}`
    }];

    for (const image of images.slice(0, 6)) {
      if (typeof image === 'string' && image.startsWith('data:image/')) {
        userContent.push({ type: 'input_image', image_url: image });
      }
    }

    const response = await fetch('https://api.openai.com/v1/responses', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${process.env.OPENAI_API_KEY}`
      },
      body: JSON.stringify({
        model: process.env.OPENAI_MODEL || 'gpt-4.1-mini',
        input: [
          { role: 'system', content: [{ type: 'input_text', text: SYSTEM_PROMPT }] },
          { role: 'user', content: userContent }
        ],
        text: {
          format: {
            type: 'json_schema',
            name: 'quotation_extraction',
            strict: true,
            schema
          }
        }
      })
    });

    const data = await response.json();
    if (!response.ok) {
      console.error(data);
      return res.status(response.status).json({ error: data?.error?.message || 'AI 분석 요청에 실패했습니다.' });
    }

    const outputText = extractOutputText(data);
    if (!outputText) return res.status(502).json({ error: 'AI가 분석 결과를 반환하지 않았습니다.' });

    const parsed = JSON.parse(outputText);
    parsed.items = (parsed.items || [])
      .filter(item => item.name && item.quantity > 0 && item.unitPrice >= 0)
      .map(item => ({
        ...item,
        quantity: Number(item.quantity),
        unitPrice: Math.round(Number(item.unitPrice)),
        amount: Math.round(Number(item.quantity) * Number(item.unitPrice))
      }));

    return res.status(200).json(parsed);
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: error.message || '서버 처리 중 오류가 발생했습니다.' });
  }
}
