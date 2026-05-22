const https = require('https');

// 服务器配置 - 请修改为您的服务器地址
const SERVER_CONFIG = {
  host: 'your_server_ip_here',
  port: '39999'
};

// 获取基础URL
function getBaseUrl() {
  return `https://${SERVER_CONFIG.host}:${SERVER_CONFIG.port}`;
}

const url = `${getBaseUrl()}/event-center/homePage/getMyTodoList?_t=${Date.now()}&pageNo=1&pageSize=1000&sort=createTime&order=descend&serialNumber=&subject=`;

const options = {
  headers: {
    'x-access-token': 'your_token_here'
  },
  rejectUnauthorized: false
};

https.get(url, options, (res) => {
  let data = '';

  res.on('data', (chunk) => {
    data += chunk;
  });

  res.on('end', () => {
    if (res.statusCode === 200) {
      // console.log(data);
      let list = JSON.parse(data).result.records;
      // console.log(list.length);

      list.forEach(item => {
        delete item.id;
        delete item.otherId;
        delete item.circulationId;
        delete item.deadline;
        delete item.status;
        delete item.createOrg;
        delete item.createUser;
        delete item.endTime;
        delete item.timeoutStatus;
        delete item.type;
        delete item.categoryName;
        delete item.statusForMe;
        delete item.typeName;
        delete item.nodeType;

        // 当remainTime为null时，赋值为协办
        if (item.remainTime === null) {
          item.remainTime = '协办';
        }

        console.log(item);

      });
    } else {
      console.error(`Status: ${res.statusCode}`);
    }
  });
}).on('error', (err) => {
  console.error(err.message);
});
