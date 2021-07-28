        <nav class="navbar navbar-expand-sm navbar-dark bg-dark small py-0">            
                  <ul class="navbar-nav mr-auto">
                        <li class="nav-item text-light">
                                <a class="text-light" href="index.asp">
                                    <i class="fa fa-home fa-lg"></i>
                                </a>
                               </li>
                  </ul>
                  <span class="navbar-text">
                        <% If Not rsGetUser.EOF then %>
                        <i class="fa fa-user fa-lg"></i> <%= rsGetUser.Fields.Item("name").Value %> <a href="logout.asp" class="ml-3">Logout</a>
                        <% end if %>
                      </span>
              </nav>