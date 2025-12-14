// 导入Supabase客户端库
const { createClient } = require('@supabase/supabase-js');

// 从环境变量读取配置（你需要在Vercel中设置）
const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_KEY;
const supabase = createClient(supabaseUrl, supabaseKey);

// 获取服务器时间（Supabase的数据库时间作为权威时间）
async function getServerTime() {
  const { data, error } = await supabase.rpc('get_server_time');
  if (!error && data) return new Date(data);
  return new Date(); // 降级方案：使用API服务器时间
}

module.exports = async (req, res) => {
  // 1. 设置CORS头，允许你的主程序调用
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // 2. 只处理GET请求，从查询参数获取卡密
  if (req.method !== 'GET') {
    return res.status(405).json({ error: '仅支持GET方法' });
  }

  const { key } = req.query;
  if (!key) {
    return res.status(400).json({ error: '缺少卡密参数: key' });
  }

  try {
    // 3. 查询数据库
    const { data, error } = await supabase
      .from('card_keys')
      .select('*')
      .eq('key', key)
      .maybeSingle(); // 使用maybeSingle，没查到返回null而不是错误

    if (error) {
      console.error('数据库查询错误:', error);
      return res.status(500).json({ error: '数据库查询失败' });
    }

    // 4. 判断卡密是否存在
    if (!data) {
      return res.status(200).json({
        valid: false,
        message: '卡密不存在'
      });
    }

    // 5. 获取服务器时间，进行逻辑判断
    const serverTime = await getServerTime();
    const validFrom = new Date(data.valid_from);
    const validTo = new Date(data.valid_to);

    // 6. 判断卡密状态（按照你之前的逻辑）
    let isValid = false;
    let message = '';
    let status = 'invalid';

    if (!data.is_active) {
      message = '卡密已停用';
      status = 'disabled';
    } else if (data.used_at) {
      message = '卡密已使用';
      status = 'used';
    } else if (validTo < serverTime) {
      message = '卡密已过期';
      status = 'expired';
    } else if (validFrom > serverTime) {
      message = '卡密未生效';
      status = 'pending';
    } else {
      isValid = true;
      message = '卡密有效';
      status = 'active';
    }

    // 7. 返回JSON，包含你要求的三个时间字段
    res.status(200).json({
      valid: isValid,
      status: status,           // 卡密状态: active/expired/used等
      message: message,
      key: data.key,
      note: data.note,
      // 你明确要求的三个核心时间字段：
      activated_at: data.valid_from,        // 卡密激活时间
      expires_at: data.valid_to,            // 卡密到期时间
      current_server_time: serverTime.toISOString(), // 当前服务器时间（供参考）
      // 其他可能用到的信息
      total_days: data.total_days,
      remaining_days: data.remaining_days
    });

  } catch (err) {
    console.error('API处理异常:', err);
    res.status(500).json({ error: '服务器内部错误' });
  }
};
