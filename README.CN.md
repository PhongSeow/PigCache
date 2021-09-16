# PigCache
#### [English Readme](https://github.com/PhongSeow/PigCache/blob/master/README.md)

PigCache 是一个轻量、多应用场景的 key-value 缓存系统，可以支持单进程的微型应用到大型的多服务器负载均衡的场景，大部分场景不需要如 memcached  或redis 这样的第三方服务支持，运行环境可支持 Windows 或 Linux 平台，不同场景的代码只需要引用不同的类库，源代码作细微的改动即可。

------

|                             类库                             |             应用场景             | 其他服务支持 | 读性能           | 写性能 | 运行平台       | 优点                                                         | 缺点                                               |
| :----------------------------------------------------------: | :------------------------------: | :----------: | ---------------- | ------ | -------------- | ------------------------------------------------------------ | -------------------------------------------------- |
| [PigKeyCacheLib](https://www.nuget.org/packages/PigKeyCacheLib/) |       支持单进程多线程程序       |    不需要    | 好               | 好     | Windows和Linux | 不需要第三方服务支持，不会增加故障点                         | 只能支持单进程                                     |
|                                                              | 支持同一用户下的多进程多线程程序 |    不需要    | 好               | 好     | Windows和Linux | 不需要第三方服务支持，不会增加故障点，可支持多进程           | 只能在同一用户会话内的程序间使用                   |
|                                                              |  支持同一主机下多进程多线程程序  |    不需要    | 好               | 较好   | Windows和Linux | 不需要第三方服务支持，不会增加故障点，可支持多用户间的进程使用 | 只能在同一台服务器上使用                           |
| [PigKeyCacheLib.SQLServer](https://www.nuget.org/packages/PigKeyCacheLib.SQLServer/) |       支持多服务器负载均衡       |  SQL Server  | 很好，可横向扩展 | 稍差   | Windows        | 这是读性能最好，而且可以横扩展的场景，与应用共用数据库连接，相当于没有增加第三方的服务支持，而且数据库高可用优于Redis。 | 写性能稍差                                         |
| [PigKeyCacheCoreLib.SQLServer](https://www.nuget.org/packages/PigKeyCacheCoreLib.SQLServer/) |       支持多服务器负载均衡       |  SQL Server  | 很好，可横向扩展 | 稍差   | Windows和Linux | 这是读性能最好，而且可以横向扩展的场景，与应用共用数据库连接，相当于没有增加第三方的服务支持，而且数据库高可用优于Redis。 | 写性能稍差                                         |
|               PigKeyCacheLib.Redis（即将推出）               |       支持多服务器负载均衡       |    Redis     | 很好，可横向扩展 | 很好   | Windows        | 这是性能最好，而且可以横扩展的场景                           | 需要Redis，这样会增加一个故障点和管理Redis的成本。 |

