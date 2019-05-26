using CommonLib.Global;
using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace CommonLib.Data.Helper
{
    public static class DataHelper
    {
        /// <summary>
        /// DataTableの指定位置から指定方向に進み空セルを見つけたら直前の位置を返す
        /// 位置は0始まり
        /// </summary>
        /// <param name="table">DataTable</param>
        /// <param name="startRow">指定位置(行)</param>
        /// <param name="startClm">指定位置(列)</param>
        /// <param name="direction">探索方向</param>
        /// <returns>テーブルの端の位置</returns>
        public static int GetPositionOfTableEdge(DataTable table, int startRow, int startClm, Constant.Direction direction = Constant.Direction.Right)
        {
            var result = 0;

            var srcRowCnt = table.Rows.Count;
            var srcClmCnt = table.Columns.Count;

            var sr = startRow;
            var sc = startClm;

            if (sr < 0) sr = 0;
            if (sc < 0) sc = 0;
            if (sr >= srcRowCnt) sr = srcRowCnt - 1;
            if (sc >= srcClmCnt) sc = srcClmCnt - 1;

            if (direction == Constant.Direction.Right)
            {
                for (var c = sc; ; c++)
                {
                    if (c >= srcClmCnt) break;

                    var value = table.Rows[sr][c].ToString();

                    if (string.IsNullOrEmpty(value)) break;

                    result++;
                }
                return startClm + result-1;
            }
            else if (direction == Constant.Direction.Down)
            {
                for (var r = sr; ; r++)
                {
                    if (r >= srcRowCnt) break;

                    var value = table.Rows[r][sc].ToString();

                    if (string.IsNullOrEmpty(value)) break;

                    result++;
                }
                return startRow + result-1;
            }
            else if (direction == Constant.Direction.Up)
            {
                // 未サポート
            }
            else
            {
                // 未サポート
            }

            return result;
        }

        /// <summary>
        /// DataTableの指定位置から指定方向に進み空セルを見つけるまで値を収集する
        /// 位置は0始まり
        /// </summary>
        /// <param name="table">DataTable</param>
        /// <param name="startRow">指定位置(行)</param>
        /// <param name="startClm">指定位置(列)</param>
        /// <param name="direction">探索方向</param>
        /// <returns>収集した文字列</returns>
        public static IEnumerable<string> CollectString(DataTable table, int startRow, int startClm, Constant.Direction direction = Constant.Direction.Right)
        {
            var result = new List<string>();

            var srcRowCnt = table.Rows.Count;
            var srcClmCnt = table.Columns.Count;

            var sr = startRow;
            var sc = startClm;

            if (sr < 0) sr = 0;
            if (sc < 0) sc = 0;
            if (sr >= srcRowCnt) sr = srcRowCnt-1;
            if (sc >= srcClmCnt) sc = srcClmCnt-1;

            if (direction == Constant.Direction.Right)
            {
                for(var c=sc; ;c++)
                {
                    if (c >= srcClmCnt) break;

                    var value = table.Rows[sr][c].ToString();

                    if (string.IsNullOrEmpty(value)) break;

                    result.Add(value);
                }
            }
            else if (direction == Constant.Direction.Down)
            {
                for (var r=sr; ; r++)
                {
                    if (r >= srcRowCnt) break;

                    var value = table.Rows[r][sc].ToString();

                    if (string.IsNullOrEmpty(value)) break;

                    result.Add(value);
                }
            }
            else if (direction == Constant.Direction.Up)
            {
                // 未サポート
            }
            else
            {
                // 未サポート
            }

            return result;
        }

        /// <summary>
        /// DataTableを指定範囲で切り出す
        /// 範囲指定は0始まり
        /// </summary>
        /// <param name="table">DataTable</param>
        /// <param name="startRow">開始行</param>
        /// <param name="startClm">開始列</param>
        /// <returns>DataTable</returns>
        public static DataTable TrimmingTable(DataTable table, int startRow, int startClm)
        {
            var endRow = GetPositionOfTableEdge(table, startRow, startClm, Constant.Direction.Down);
            var endClm = GetPositionOfTableEdge(table, startRow, startClm, Constant.Direction.Right);

            return TrimmingTable(table, startRow, startClm, endRow, endClm);
        }

        /// <summary>
        /// DataTableを指定範囲で切り出す
        /// 範囲指定は0始まり
        /// </summary>
        /// <param name="table">DataTable</param>
        /// <param name="range">範囲指定(開始行列,終了行列の順)</param>
        /// <returns>DataTable</returns>
        public static DataTable TrimmingTable(DataTable table, int[] range) => TrimmingTable(table, range[0], range[1], range[2], range[3]);

        /// <summary>
        /// DataTableを指定範囲で切り出す
        /// 範囲指定は0始まり
        /// </summary>
        /// <param name="table">DataTable</param>
        /// <param name="start">範囲指定(開始行列)</param>
        /// <param name="end">範囲指定(終了行列)</param>
        /// <returns>DataTable</returns>
        public static DataTable TrimmingTable(DataTable table, int[] start, int[] end)
        {
            return TrimmingTable(table, start[0], start[1], end[0], end[1]);
        }

        /// <summary>
        /// DataTableを指定範囲で切り出す
        /// 範囲指定は0始まり
        /// </summary>
        /// <param name="table">DataTable</param>
        /// <param name="startRow">開始行</param>
        /// <param name="startClm">開始列</param>
        /// <param name="endRow">終了行</param>
        /// <param name="endClm">終了列</param>
        /// <returns>DataTable</returns>
        public static DataTable TrimmingTable(DataTable table, int startRow, int startClm, int endRow, int endClm)
        {
            var srcRowCnt = table.Rows.Count;
            var srcClmCnt = table.Columns.Count;

            var sr = Math.Min(startRow, endRow);
            var sc = Math.Min(startClm, endClm);

            var er = Math.Max(startRow, endRow);
            var ec = Math.Max(startClm, endClm);

            var rowLen = (er - sr)+1;
            var clmLen = (ec - sc)+1;

            if ( (rowLen <= 0) || (clmLen <= 0) ||
                 (srcRowCnt < rowLen) || (srcClmCnt < clmLen) )
            {
                return table.Copy();
            }

            var result = table.Copy();

            int c;  // ループカウンタ(常に++)
            int len;// ループ数
            int p;  // ポインタ(スキップしたら++)
            int d;  // 削除位置(スキップ前=0, スキップ後=e+1)
            int s;  // スキップ開始(削除したら--)
            int e;  // スキップ終了(削除したら--)

            len = srcClmCnt;

            // 列を削除
            for (c=0, p=0, d=0, s=sc, e=ec; c<len; c++)
            {
                // スキップ判定(s <= p <= e)
                if (s <= p && p <= e)
                {
                    p++;
                    d = e+1;
                    continue;
                }

                // 削除
                result.Columns.RemoveAt(d);
                s--;
                e--;
            }

            len = srcRowCnt;

            // 行を削除
            for (c = 0, p = 0, d = 0, s = sr, e = er; c < len; c++)
            {
                // スキップ判定(s <= p <= e)
                if (s <= p && p <= e)
                {
                    p++;
                    d = e + 1;
                    continue;
                }

                // 削除
                result.Rows.RemoveAt(d);
                s--;
                e--;
            }

            return result;
        }

        /// <summary>
        /// プロパティ情報を返す
        /// </summary>
        /// <param name="type">型</param>
        /// <param name="propertyName">
        /// プロパティ名
        /// .で階層的に指定可能(Ex."Position.X")
        /// </param>
        /// <returns>プロパティ情報</returns>
        public static PropertyInfo GetPropertyInfo(Type type, string propertyName)
        {
            PropertyInfo result = null;

            // 階層的に指定されている場合があるため"."で分割
            var level = propertyName.Split('.');

            // 階層数分ループ
            for (var i = 0; i < level.Length; i++)
            {
                // 1階層目のプロパティ情報を取得
                result = type.GetProperty(level[i]);

                // 取得成功 かつ 2階層以上ある場合は再帰的に取得
                if (result != null && level.Length > 1)
                {
                    // 再帰
                    result = GetPropertyInfo(
                                Type.GetType(result.PropertyType.FullName),
                                propertyName.Remove(0, level[i].Length + 1));   // すでに取得したプロパティ名は削除。+1は"."

                    if (result != null)
                    {
                        // 子階層の取得が成功したのでループを抜ける
                        break;
                    }
                }
                else
                {
                    return result;
                }
            }
            return result;
        }
    }
}
