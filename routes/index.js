const express = require('express');
const router = express.Router();
const util = require('../util/tools');
/* GET home page. */
router.get('/', function (req, res, next) {
  res.render('index', { title: 'Express' });
});
// 导出excal
router.get('/exportExcal', async (req, res, next) => {
  let _rows, _columns, _sheet_name, _style, resource, resourceChild, _data, _row;
  const list = [
    {
      a: 'aaaa',
      b: 'bbb'
    },
    {
      a: 'cccc',
      b: 'ddddd'
    }
  ]
  _sheet_name = '导出excal样例';
  _rows = [];
  _style = {
    border: {
      top: {
        style: 'thin'
      },
      left: {
        style: 'thin'
      },
      bottom: {
        style: 'thin'
      },
      right: {
        style: 'thin'
      }
    },
    alignment: {
      horizontal: 'left'
    }
  };
  _columns = [
    {
      header: '单位',
      key: 'orgName',
      width: 25,
      style: _style
    }, {
      header: '账号',
      key: 'accountId',
      width: 25,
      style: _style
    }
  ];
  for (let i = 0; i < list.length; i++) {
    resource = list[i];
    _row = {
      orgName: resource['a'],
      accountId: resource['b']
    }
    _rows.push(_row);
  }
  _data = {
    rows: _rows,
    columns: _columns,
    sheet_name: _sheet_name,
    path: '/tmp/' + _sheet_name + '.xlsx'
  };
  
 let path = await util.createExcel(_data);
 res.download(path)
})
module.exports = router;
