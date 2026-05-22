const https = require('https');

const url = "https://59.215.230.158:39999/event-center/homePage/getMyTodoList?_t=1775123755&pageNo=1&pageSize=1000&sort=createTime&order=descend&serialNumber=&subject=";

const options = {
  headers: {
    'x-access-token': 'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJkZXB0SWQiOiIxNTc1MDY3Nzk5NjMyNTA2OTcxIiwiaWQiOiIxODcyNDY0NzE0MTE3OTgwMTYyIiwiZXhwIjoxNzc1MTE0MjAwLCJ1c2VybmFtZSI6IkR6ZnhtZ2pkIn0.4TLllNP_2L9lDa4qKdLCX9786H0SzN0RVrBv6W0fBzE'
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